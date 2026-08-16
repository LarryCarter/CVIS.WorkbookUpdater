using System.IO;
using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CVIS.WorkbookUpdater.Services;

internal static class WorkbookCompatibilityService
{
    private const string ValidationSheetName = "CVIS Validation Lists";

    public static XLWorkbook Open(string workbookPath)
    {
        var tempPath = Path.Combine(
            Path.GetTempPath(),
            $"cvis-compatible-{Guid.NewGuid():N}.xlsx");
        File.Copy(workbookPath, tempPath);
        try
        {
            NormalizeLongInlineValidations(tempPath);
            var content = File.ReadAllBytes(tempPath);
            return new XLWorkbook(new MemoryStream(content, writable: false));
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    internal static int NormalizeLongInlineValidations(string workbookPath)
    {
        using var document = SpreadsheetDocument.Open(workbookPath, true);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidOperationException("Workbook has no workbook part.");
        var longValidations = workbookPart.WorksheetParts
            .SelectMany(part => part.Worksheet.Descendants<DataValidation>()
                .Where(validation => validation.Type?.Value == DataValidationValues.List)
                .Select(validation => (Part: part, Validation: validation)))
            .Where(item => item.Validation.Formula1?.Text?.Length > 255)
            .ToList();
        if (longValidations.Count == 0)
        {
            return 0;
        }

        var helperPart = GetOrCreateHelperSheet(workbookPart);
        var sheetData = helperPart.Worksheet.GetFirstChild<SheetData>()
            ?? helperPart.Worksheet.AppendChild(new SheetData());
        var columnIndex = 1;
        foreach (var item in longValidations)
        {
            var values = ParseInlineList(item.Validation.Formula1!.Text!);
            if (values.Count == 0)
            {
                throw new InvalidOperationException("A long data-validation list could not be parsed.");
            }

            for (var rowIndex = 1; rowIndex <= values.Count; rowIndex++)
            {
                var row = sheetData.Elements<Row>().FirstOrDefault(existing => existing.RowIndex?.Value == (uint)rowIndex);
                if (row is null)
                {
                    row = new Row { RowIndex = (uint)rowIndex };
                    sheetData.Append(row);
                }

                var reference = $"{ColumnName(columnIndex)}{rowIndex}";
                var cell = new Cell
                {
                    CellReference = reference,
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text(values[rowIndex - 1]))
                };
                row.Append(cell);
            }

            item.Validation.Formula1!.Text =
                $"'{ValidationSheetName}'!${ColumnName(columnIndex)}$1:${ColumnName(columnIndex)}${values.Count}";
            item.Part.Worksheet.Save();
            columnIndex++;
        }

        helperPart.Worksheet.Save();
        workbookPart.Workbook.Save();
        return longValidations.Count;
    }

    private static WorksheetPart GetOrCreateHelperSheet(WorkbookPart workbookPart)
    {
        var sheets = workbookPart.Workbook.Sheets ?? workbookPart.Workbook.AppendChild(new Sheets());
        var existing = sheets.Elements<Sheet>()
            .FirstOrDefault(sheet => sheet.Name?.Value == ValidationSheetName);
        if (existing?.Id?.Value is string relationshipId)
        {
            return (WorksheetPart)workbookPart.GetPartById(relationshipId);
        }

        var part = workbookPart.AddNewPart<WorksheetPart>();
        part.Worksheet = new Worksheet(new SheetData());
        var nextId = sheets.Elements<Sheet>().Select(sheet => sheet.SheetId?.Value ?? 0U).DefaultIfEmpty().Max() + 1;
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(part),
            SheetId = nextId,
            Name = ValidationSheetName,
            State = SheetStateValues.VeryHidden
        });
        return part;
    }

    private static IReadOnlyList<string> ParseInlineList(string formula)
    {
        var text = formula.Trim();
        if (text.Length >= 2 && text[0] == '"' && text[^1] == '"')
        {
            text = text[1..^1];
        }

        var values = new List<string>();
        var current = new StringBuilder();
        var quoted = false;
        for (var index = 0; index < text.Length; index++)
        {
            var character = text[index];
            if (character == '"')
            {
                if (quoted && index + 1 < text.Length && text[index + 1] == '"')
                {
                    current.Append('"');
                    index++;
                }
                else
                {
                    quoted = !quoted;
                }
            }
            else if (character == ',' && !quoted)
            {
                values.Add(current.ToString().Trim());
                current.Clear();
            }
            else
            {
                current.Append(character);
            }
        }

        values.Add(current.ToString().Trim());
        return values.Where(value => !string.IsNullOrWhiteSpace(value)).ToList();
    }

    private static string ColumnName(int index)
    {
        var result = string.Empty;
        while (index > 0)
        {
            index--;
            result = (char)('A' + index % 26) + result;
            index /= 26;
        }

        return result;
    }
}
