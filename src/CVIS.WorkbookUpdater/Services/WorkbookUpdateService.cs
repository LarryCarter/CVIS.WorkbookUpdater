using ClosedXML.Excel;
using CVIS.WorkbookUpdater.Models;

namespace CVIS.WorkbookUpdater.Services;

public sealed class WorkbookUpdateService
{
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

        if (IsFileLocked(workbookPath))
        {
            throw new IOException("The workbook is open or locked. Close it in Excel, wait for OneDrive sync, then retry.");
        }

        using var workbook = new XLWorkbook(workbookPath);
        var intake = EnsureSheetWithHeaders(workbook, "Data Intake", DataIntakeHeaders);
        var changeLog = EnsureSheetWithHeaders(workbook, "Change Log", ChangeLogHeaders);

        var nextIntakeRow = intake.LastRowUsed()?.RowNumber() + 1 ?? 2;
        var nextChangeRow = changeLog.LastRowUsed()?.RowNumber() + 1 ?? 2;

        for (var i = 0; i < updates.Count; i++)
        {
            var update = updates[i];
            if (string.IsNullOrWhiteSpace(update.IntakeId))
            {
                update.IntakeId = $"IN-{DateTime.Now:yyyyMMddHHmmss}-{i + 1:000}";
            }

            WriteIntakeRow(intake, nextIntakeRow++, update);
            WriteChangeLogRow(changeLog, nextChangeRow++, update);
        }

        if (!dryRun)
        {
            workbook.SaveAs(workbookPath);
        }

        return updates.Count;
    }

    public IReadOnlyList<MappingRow> ReadMappings(string workbookPath)
    {
        using var workbook = new XLWorkbook(workbookPath);
        if (!workbook.TryGetWorksheet("Import Mapping", out var sheet))
        {
            return Array.Empty<MappingRow>();
        }

        var headerMap = ReadHeaderMap(sheet);
        var rows = new List<MappingRow>();
        foreach (var row in sheet.RowsUsed().Skip(1))
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

    private static IXLWorksheet EnsureSheetWithHeaders(XLWorkbook workbook, string sheetName, string[] headers)
    {
        var sheet = workbook.TryGetWorksheet(sheetName, out var existing)
            ? existing
            : workbook.Worksheets.Add(sheetName);

        for (var col = 0; col < headers.Length; col++)
        {
            if (sheet.Cell(1, col + 1).IsEmpty())
            {
                sheet.Cell(1, col + 1).Value = headers[col];
            }
        }

        return sheet;
    }

    private static Dictionary<string, int> ReadHeaderMap(IXLWorksheet sheet)
    {
        return sheet.Row(1)
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
            sheet.Cell(row, i + 1).Value = values[i]?.ToString() ?? string.Empty;
        }
    }

    private static void WriteChangeLogRow(IXLWorksheet sheet, int row, IntakeUpdate update)
    {
        object?[] values =
        [
            $"CHG-{DateTime.Now:yyyyMMddHHmmss}-{row - 1:000}",
            DateTime.Now,
            update.SubmittedBy,
            update.TargetTable,
            $"{update.RecordKeyField}={update.RecordKeyValue}",
            update.FieldToUpdate,
            string.Empty,
            update.NewValue,
            string.IsNullOrWhiteSpace(update.SourceFile) ? "WPF Manual Entry" : update.SourceFile,
            update.ReviewStatus,
            update.Notes
        ];

        for (var i = 0; i < values.Length; i++)
        {
            sheet.Cell(row, i + 1).Value = values[i]?.ToString() ?? string.Empty;
        }
    }
}
