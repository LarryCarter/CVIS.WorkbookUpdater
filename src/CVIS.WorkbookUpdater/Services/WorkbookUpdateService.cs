using System.IO;
using ClosedXML.Excel;
using CVIS.WorkbookUpdater.Models;

namespace CVIS.WorkbookUpdater.Services;

public sealed class WorkbookUpdateService
{
    private readonly IWorkbookSourceProvider _sourceProvider;

    public WorkbookUpdateService()
        : this(new WorkbookSourceProvider())
    {
    }

    public WorkbookUpdateService(IWorkbookSourceProvider sourceProvider)
    {
        _sourceProvider = sourceProvider;
    }

    private static readonly string[] DataIntakeHeaders =
    [
        "Intake ID", "Submitted Date", "Submitted By", "Update Type", "Target Table",
        "Record Key Field", "Record Key Value", "Field To Update", "New Value",
        "Effective Date", "Source File", "Source Row", "Review Status",
        "Reviewed By", "Applied Date", "Notes"
    ];

    private static readonly string[] ChangeLogHeaders =
    [
        "Change ID", "Timestamp", "Changed By", "Target Table", "Record Key",
        "Field", "Old Value", "New Value", "Source", "Review Status", "Notes"
    ];

    public bool IsFileLocked(string workbookPath)
    {
        try
        {
            using var stream = File.Open(workbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false;
        }
        catch (IOException)
        {
            return true;
        }
    }

    public int AppendIntakeUpdates(string workbookPath, IReadOnlyList<IntakeUpdate> updates, bool dryRun)
    {
        if (updates.Count == 0)
        {
            return 0;
        }

        using var source = _sourceProvider.Open(workbookPath, writable: !dryRun);
        var localPath = source.LocalPath;
        if (IsFileLocked(localPath))
        {
            throw new IOException("The workbook is open or locked. Close it in Excel, wait for OneDrive sync, then retry.");
        }

        using var workbook = WorkbookCompatibilityService.Open(localPath);
        AppendRows(workbook, updates, includeChangeLog: true);

        if (!dryRun)
        {
            var backupPath = SaveAtomically(workbook, localPath);
            _ = source.PreserveBackup(backupPath);
            source.Commit();
        }

        return updates.Count;
    }

    public IReadOnlyList<UpdatePreview> PreviewUpdates(
        string workbookPath,
        IReadOnlyList<IntakeUpdate> updates)
    {
        using var source = _sourceProvider.Open(workbookPath, writable: false);
        using var workbook = WorkbookCompatibilityService.Open(source.LocalPath);
        var overlapping = updates
            .GroupBy(UpdateIdentity, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Select(update => update.NewValue.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1)
            .Select(group => group.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        return updates.Select(update => PreviewUpdate(workbook, update, overlapping.Contains(UpdateIdentity(update)))).ToList();
    }

    public WorkbookSaveResult ApplyUpdates(
        string workbookPath,
        IReadOnlyList<IntakeUpdate> updates,
        string reviewedBy)
    {
        if (updates.Count == 0)
        {
            return new WorkbookSaveResult();
        }

        using var source = _sourceProvider.Open(workbookPath, writable: true);
        var localPath = source.LocalPath;
        if (IsFileLocked(localPath))
        {
            throw new IOException("The workbook is open or locked. Close it in Excel, wait for OneDrive sync, then retry.");
        }

        using var workbook = WorkbookCompatibilityService.Open(localPath);
        var overlapping = updates
            .GroupBy(UpdateIdentity, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Select(update => update.NewValue.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1)
            .Select(group => group.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var previews = updates
            .Select(update => PreviewUpdate(workbook, update, overlapping.Contains(UpdateIdentity(update))))
            .ToList();
        var blocked = previews.Where(preview => !preview.CanApply).ToList();
        if (blocked.Count > 0)
        {
            var reasons = string.Join(Environment.NewLine, blocked.Take(10).Select(preview =>
                $"{preview.Update.TargetTable}/{preview.Update.RecordKeyValue}/{preview.Update.FieldToUpdate}: {preview.Result}"));
            throw new InvalidOperationException($"{blocked.Count} update(s) cannot be applied:{Environment.NewLine}{reasons}");
        }

        var changeLog = EnsureSheetWithHeaders(workbook, "Change Log", ChangeLogHeaders, out _);
        var nextChangeRow = NextDataRow(changeLog);
        foreach (var preview in previews)
        {
            var update = preview.Update;
            var target = ResolveTarget(workbook, update)
                ?? throw new InvalidOperationException($"Target disappeared while applying {UpdateIdentity(update)}.");
            SetTargetCellValue(target.Cell, update.NewValue);
            update.ReviewStatus = "Applied";
            update.ReviewedBy = reviewedBy;
            update.AppliedDate = DateTime.Now;
            WriteChangeLogRow(changeLog, nextChangeRow++, update, preview.CurrentValue);
        }

        AppendRows(workbook, updates, includeChangeLog: false);
        var backupPath = SaveAtomically(workbook, localPath);
        var preservedBackupPath = source.PreserveBackup(backupPath);
        source.Commit();
        return new WorkbookSaveResult { UpdateCount = updates.Count, BackupPath = preservedBackupPath };
    }

    public IReadOnlyList<MappingRow> ReadMappings(string workbookPath)
    {
        using var source = _sourceProvider.Open(workbookPath, writable: false);
        using var workbook = WorkbookCompatibilityService.Open(source.LocalPath);
        if (!workbook.TryGetWorksheet("Import Mapping", out var sheet))
        {
            return Array.Empty<MappingRow>();
        }

        var headerRow = FindHeaderRow(sheet, "Mapping Profile", "Source Field", "Target Field");
        if (headerRow is null)
        {
            return Array.Empty<MappingRow>();
        }

        var headerMap = ReadHeaderMap(sheet, headerRow.Value);
        var rows = new List<MappingRow>();
        foreach (var row in sheet.RowsUsed().Where(row => row.RowNumber() > headerRow.Value))
        {
            var mappingProfile = GetCell(row, headerMap, "Mapping Profile");
            var sourceField = GetCell(row, headerMap, "Source Field");
            var targetField = GetCell(row, headerMap, "Target Field");
            if (string.IsNullOrWhiteSpace(mappingProfile) || string.IsNullOrWhiteSpace(sourceField))
            {
                continue;
            }

            rows.Add(new MappingRow
            {
                MappingProfile = mappingProfile,
                SourceType = GetCell(row, headerMap, "Source Type"),
                TargetTable = GetCell(row, headerMap, "Target Table"),
                SourceField = sourceField,
                TargetField = targetField,
                Required = IsYes(GetCell(row, headerMap, "Required?")),
                MatchKey = IsYes(GetCell(row, headerMap, "Match Key?")),
                TransformRule = GetCell(row, headerMap, "Transform Rule"),
                DefaultValue = GetCell(row, headerMap, "Default Value"),
                Notes = GetCell(row, headerMap, "Notes")
            });
        }

        return rows;
    }

    private static UpdatePreview PreviewUpdate(XLWorkbook workbook, IntakeUpdate update, bool overlaps)
    {
        if (overlaps)
        {
            return Preview(update, false, false, string.Empty, "Conflicting overlapping updates require review.");
        }

        var target = ResolveTarget(workbook, update);
        if (target is null)
        {
            return Preview(update, false, false, string.Empty, "Target sheet, key, record, or field was not found.");
        }

        var currentValue = target.Value.Cell.GetFormattedString().Trim();
        if (string.IsNullOrWhiteSpace(update.NewValue) && !string.IsNullOrWhiteSpace(currentValue))
        {
            return Preview(update, true, false, currentValue, "A blank value cannot overwrite an existing value.");
        }

        if (string.Equals(currentValue, update.NewValue.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            return Preview(update, true, true, currentValue, "No change; value already matches.");
        }

        return Preview(update, true, true, currentValue, "Ready to apply.");
    }

    private static UpdatePreview Preview(
        IntakeUpdate update,
        bool targetFound,
        bool canApply,
        string currentValue,
        string result) => new()
    {
        Update = update,
        TargetFound = targetFound,
        CanApply = canApply,
        CurrentValue = currentValue,
        Result = result
    };

    private static (IXLCell Cell, int HeaderRow)? ResolveTarget(XLWorkbook workbook, IntakeUpdate update)
    {
        if (!workbook.TryGetWorksheet(update.TargetTable, out var sheet))
        {
            return null;
        }

        var headerRow = FindHeaderRow(sheet, update.RecordKeyField, update.FieldToUpdate);
        if (headerRow is null)
        {
            return null;
        }

        var headers = ReadHeaderMap(sheet, headerRow.Value);
        if (!headers.TryGetValue(update.RecordKeyField, out var keyColumn)
            || !headers.TryGetValue(update.FieldToUpdate, out var targetColumn))
        {
            return null;
        }

        var targetRow = sheet.RowsUsed()
            .Where(row => row.RowNumber() > headerRow.Value)
            .FirstOrDefault(row => row.Cell(keyColumn).GetFormattedString().Trim()
                .Equals(update.RecordKeyValue.Trim(), StringComparison.OrdinalIgnoreCase));
        return targetRow is null ? null : (targetRow.Cell(targetColumn), headerRow.Value);
    }

    private static string UpdateIdentity(IntakeUpdate update) =>
        $"{update.TargetTable}\u001f{update.RecordKeyField}\u001f{update.RecordKeyValue}\u001f{update.FieldToUpdate}";

    private static void AppendRows(XLWorkbook workbook, IReadOnlyList<IntakeUpdate> updates, bool includeChangeLog)
    {
        var intake = EnsureSheetWithHeaders(workbook, "Data Intake", DataIntakeHeaders, out _);
        var nextIntakeRow = NextDataRow(intake);
        IXLWorksheet? changeLog = null;
        var nextChangeRow = 0;
        if (includeChangeLog)
        {
            changeLog = EnsureSheetWithHeaders(workbook, "Change Log", ChangeLogHeaders, out _);
            nextChangeRow = NextDataRow(changeLog);
        }

        for (var i = 0; i < updates.Count; i++)
        {
            var update = updates[i];
            if (string.IsNullOrWhiteSpace(update.IntakeId))
            {
                update.IntakeId = $"IN-{DateTime.Now:yyyyMMddHHmmssfff}-{i + 1:000}";
            }

            CopyPriorRowStyle(intake, nextIntakeRow);
            WriteIntakeRow(intake, nextIntakeRow++, update);
            if (changeLog is not null)
            {
                CopyPriorRowStyle(changeLog, nextChangeRow);
                WriteChangeLogRow(changeLog, nextChangeRow++, update, string.Empty);
            }
        }
    }

    private static int NextDataRow(IXLWorksheet sheet) => sheet.LastRowUsed()?.RowNumber() + 1 ?? 2;

    private static void CopyPriorRowStyle(IXLWorksheet sheet, int row)
    {
        if (row <= 1)
        {
            return;
        }

        var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 1;
        sheet.Range(row - 1, 1, row - 1, lastColumn).CopyTo(sheet.Cell(row, 1));
        sheet.Range(row, 1, row, lastColumn).Clear(XLClearOptions.Contents);
    }

    private static string SaveAtomically(XLWorkbook workbook, string workbookPath)
    {
        var fullPath = Path.GetFullPath(workbookPath);
        var directory = Path.GetDirectoryName(fullPath)
            ?? throw new InvalidOperationException("Workbook path has no parent directory.");
        var backupDirectory = Path.Combine(directory, ".cvis-backups");
        Directory.CreateDirectory(backupDirectory);
        var backupPath = Path.Combine(
            backupDirectory,
            $"{Path.GetFileNameWithoutExtension(fullPath)}.{DateTime.Now:yyyyMMddHHmmssfff}.xlsx");
        var tempPath = Path.Combine(directory, $".{Path.GetFileName(fullPath)}.{Guid.NewGuid():N}.tmp.xlsx");

        try
        {
            workbook.SaveAs(tempPath);
            using (var verification = new XLWorkbook(tempPath))
            {
                _ = verification.Worksheets.Count;
            }

            File.Replace(tempPath, fullPath, backupPath, ignoreMetadataErrors: true);
            try
            {
                using var verification = new XLWorkbook(fullPath);
                _ = verification.Worksheets.Count;
            }
            catch
            {
                File.Copy(backupPath, fullPath, overwrite: true);
                throw;
            }

            return backupPath;
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }

    private static void SetTargetCellValue(IXLCell cell, string value)
    {
        var trimmed = value.Trim();
        if (cell.DataType == XLDataType.DateTime && DateTime.TryParse(trimmed, out var date))
        {
            cell.Value = date;
        }
        else if (cell.DataType == XLDataType.Number
            && double.TryParse(trimmed, System.Globalization.NumberStyles.Number,
                System.Globalization.CultureInfo.InvariantCulture, out var number))
        {
            cell.Value = number;
        }
        else
        {
            cell.Value = trimmed;
        }
    }

    private static IXLWorksheet EnsureSheetWithHeaders(
        XLWorkbook workbook,
        string sheetName,
        string[] headers,
        out int headerRow)
    {
        var sheet = workbook.TryGetWorksheet(sheetName, out var existing)
            ? existing
            : workbook.Worksheets.Add(sheetName);

        headerRow = FindHeaderRow(sheet, headers) ?? 1;
        for (var col = 0; col < headers.Length; col++)
        {
            if (sheet.Cell(headerRow, col + 1).IsEmpty())
            {
                sheet.Cell(headerRow, col + 1).Value = headers[col];
            }
        }

        return sheet;
    }

    private static int? FindHeaderRow(IXLWorksheet sheet, params string[] requiredHeaders)
    {
        var required = requiredHeaders.ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var row in sheet.RowsUsed())
        {
            var values = row.CellsUsed()
                .Select(cell => cell.GetString().Trim())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (required.IsSubsetOf(values))
            {
                return row.RowNumber();
            }
        }

        return null;
    }

    private static Dictionary<string, int> ReadHeaderMap(IXLWorksheet sheet, int headerRow)
    {
        return sheet.Row(headerRow)
            .CellsUsed()
            .ToDictionary(
                cell => cell.GetString().Trim(),
                cell => cell.Address.ColumnNumber,
                StringComparer.OrdinalIgnoreCase);
    }

    private static string GetCell(IXLRow row, Dictionary<string, int> headerMap, string header)
    {
        return headerMap.TryGetValue(header, out var col)
            ? row.Cell(col).GetString().Trim()
            : string.Empty;
    }

    private static bool IsYes(string value)
    {
        return value.Equals("yes", StringComparison.OrdinalIgnoreCase)
            || value.Equals("y", StringComparison.OrdinalIgnoreCase)
            || value.Equals("true", StringComparison.OrdinalIgnoreCase);
    }

    private static void WriteIntakeRow(IXLWorksheet sheet, int row, IntakeUpdate update)
    {
        object?[] values =
        [
            update.IntakeId,
            update.SubmittedDate,
            update.SubmittedBy,
            update.UpdateType,
            update.TargetTable,
            update.RecordKeyField,
            update.RecordKeyValue,
            update.FieldToUpdate,
            update.NewValue,
            update.EffectiveDate,
            update.SourceFile,
            update.SourceRow,
            update.ReviewStatus,
            update.ReviewedBy,
            update.AppliedDate,
            update.Notes
        ];

        for (var i = 0; i < values.Length; i++)
        {
            SetCellValue(sheet.Cell(row, i + 1), values[i]);
        }

        sheet.Cell(row, 2).Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
        sheet.Cell(row, 10).Style.DateFormat.Format = "yyyy-mm-dd";
        sheet.Cell(row, 15).Style.DateFormat.Format = "yyyy-mm-dd";
    }

    private static void WriteChangeLogRow(IXLWorksheet sheet, int row, IntakeUpdate update, string oldValue)
    {
        object?[] values =
        [
            $"CHG-{DateTime.Now:yyyyMMddHHmmss}-{row - 1:000}",
            DateTime.Now,
            update.SubmittedBy,
            update.TargetTable,
            $"{update.RecordKeyField}={update.RecordKeyValue}",
            update.FieldToUpdate,
            oldValue,
            update.NewValue,
            string.IsNullOrWhiteSpace(update.SourceFile) ? "WPF Manual Entry" : update.SourceFile,
            update.ReviewStatus,
            update.Notes
        ];

        for (var i = 0; i < values.Length; i++)
        {
            SetCellValue(sheet.Cell(row, i + 1), values[i]);
        }

        sheet.Cell(row, 2).Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
    }

    private static void SetCellValue(IXLCell cell, object? value)
    {
        switch (value)
        {
            case null:
                cell.Clear(XLClearOptions.Contents);
                break;
            case DateTime dateTime:
                cell.Value = dateTime;
                break;
            case bool boolean:
                cell.Value = boolean;
                break;
            case int integer:
                cell.Value = integer;
                break;
            case long longInteger:
                cell.Value = longInteger;
                break;
            case double doubleNumber:
                cell.Value = doubleNumber;
                break;
            case decimal decimalNumber:
                cell.Value = decimalNumber;
                break;
            default:
                cell.Value = value.ToString() ?? string.Empty;
                break;
        }
    }
}
