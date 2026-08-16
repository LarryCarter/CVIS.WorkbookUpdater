using ClosedXML.Excel;
using CVIS.WorkbookUpdater.Models;
using CVIS.WorkbookUpdater.Services;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CVIS.WorkbookUpdater.Tests;

public sealed class WorkbookUpdateServiceTests : IDisposable
{
    private readonly string _tempDirectory = Path.Combine(
        Path.GetTempPath(),
        $"cvis-workbook-updater-tests-{Guid.NewGuid():N}");

    public WorkbookUpdateServiceTests()
    {
        Directory.CreateDirectory(_tempDirectory);
    }

    [Fact]
    public void ReadMappings_FindsStyledHeaderOnRowFour()
    {
        var path = Path.Combine(_tempDirectory, "mapping.xlsx");
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Import Mapping");
            sheet.Cell("A1").Value = "Import Mapping";
            sheet.Cell("A2").Value = "Presentation text that is not part of the table.";
            sheet.Cell("A4").InsertData(new[]
            {
                new[]
                {
                    "Mapping Profile", "Source Type", "Target Table", "Source Field", "Target Field",
                    "Required?", "Match Key?", "Transform Rule", "Default Value", "Notes"
                },
                new[]
                {
                    "ApplicationOwnerIntake", "CSV/YAML", "Application Inventory", "app_id", "App ID",
                    "Yes", "Yes", "Uppercase", "", "Preferred key"
                }
            });
            workbook.SaveAs(path);
        }

        var mappings = new WorkbookUpdateService().ReadMappings(path);

        var mapping = Assert.Single(mappings);
        Assert.Equal("ApplicationOwnerIntake", mapping.MappingProfile);
        Assert.Equal("app_id", mapping.SourceField);
        Assert.True(mapping.Required);
        Assert.True(mapping.MatchKey);
    }

    [Fact]
    public void AppendIntakeUpdates_PreservesPresentationRowsAndAppendsBelowRowFourHeader()
    {
        var path = Path.Combine(_tempDirectory, "intake.xlsx");
        using (var workbook = new XLWorkbook())
        {
            var intake = workbook.AddWorksheet("Data Intake");
            intake.Cell("A1").Value = "Data Intake Queue";
            intake.Cell("A2").Value = "Presentation text";
            intake.Cell("A4").InsertData(new[] { new[]
            {
                "Intake ID", "Submitted Date", "Submitted By", "Update Type", "Target Table",
                "Record Key Field", "Record Key Value", "Field To Update", "New Value",
                "Effective Date", "Source File", "Source Row", "Review Status",
                "Reviewed By", "Applied Date", "Notes"
            }});

            var log = workbook.AddWorksheet("Change Log");
            log.Cell("A1").Value = "Change Log";
            log.Cell("A4").InsertData(new[] { new[]
            {
                "Change ID", "Timestamp", "Changed By", "Target Table", "Record Key",
                "Field", "Old Value", "New Value", "Source", "Review Status", "Notes"
            }});
            workbook.SaveAs(path);
        }

        var update = new IntakeUpdate
        {
            SubmittedBy = "Test User",
            RecordKeyValue = "APP-001",
            FieldToUpdate = "Owner",
            NewValue = "CVIS Team"
        };

        var count = new WorkbookUpdateService().AppendIntakeUpdates(path, [update], dryRun: false);

        Assert.Equal(1, count);
        using var result = new XLWorkbook(path);
        var intakeResult = result.Worksheet("Data Intake");
        Assert.Equal("Data Intake Queue", intakeResult.Cell("A1").GetString());
        Assert.Equal("Intake ID", intakeResult.Cell("A4").GetString());
        Assert.StartsWith("IN-", intakeResult.Cell("A5").GetString());
        Assert.Equal("Test User", intakeResult.Cell("C5").GetString());
        Assert.Equal("Owner", intakeResult.Cell("H5").GetString());
        Assert.Equal(XLDataType.DateTime, intakeResult.Cell("B5").DataType);
        Assert.Equal(XLDataType.DateTime, intakeResult.Cell("J5").DataType);

        var logResult = result.Worksheet("Change Log");
        Assert.Equal("Change ID", logResult.Cell("A4").GetString());
        Assert.StartsWith("CHG-", logResult.Cell("A5").GetString());
        Assert.Equal(XLDataType.DateTime, logResult.Cell("B5").DataType);
    }

    [Fact]
    public void ApplyUpdates_ChangesTargetAndCreatesVerifiedBackupAndAuditRows()
    {
        var path = CreateApplyWorkbook();
        var update = new IntakeUpdate
        {
            SubmittedBy = "Test User",
            TargetTable = "Application Inventory",
            RecordKeyField = "App ID",
            RecordKeyValue = "APP-001",
            FieldToUpdate = "Owner",
            NewValue = "New Owner"
        };
        var service = new WorkbookUpdateService();

        var preview = Assert.Single(service.PreviewUpdates(path, [update]));
        Assert.True(preview.CanApply);
        Assert.Equal("Old Owner", preview.CurrentValue);

        var result = service.ApplyUpdates(path, [update], "Reviewer");

        Assert.Equal(1, result.UpdateCount);
        Assert.True(File.Exists(result.BackupPath));
        using (var updated = new XLWorkbook(path))
        {
            Assert.Equal("New Owner", updated.Worksheet("Application Inventory").Cell("C5").GetString());
            Assert.Equal("Applied", updated.Worksheet("Data Intake").Cell("M5").GetString());
            Assert.Equal("Reviewer", updated.Worksheet("Data Intake").Cell("N5").GetString());
            Assert.Equal("Old Owner", updated.Worksheet("Change Log").Cell("G5").GetString());
            Assert.Equal("New Owner", updated.Worksheet("Change Log").Cell("H5").GetString());
        }

        using var backup = new XLWorkbook(result.BackupPath);
        Assert.Equal("Old Owner", backup.Worksheet("Application Inventory").Cell("C5").GetString());
    }

    [Fact]
    public void PreviewUpdates_BlocksConflictingOverlapsAndBlankOverwrite()
    {
        var path = CreateApplyWorkbook();
        var first = OwnerUpdate("First Owner");
        var second = OwnerUpdate("Second Owner");
        var blank = new IntakeUpdate
        {
            TargetTable = "Application Inventory",
            RecordKeyField = "App ID",
            RecordKeyValue = "APP-001",
            FieldToUpdate = "Status",
            NewValue = ""
        };

        var previews = new WorkbookUpdateService().PreviewUpdates(path, [first, second, blank]);

        Assert.False(previews[0].CanApply);
        Assert.Contains("overlapping", previews[0].Result, StringComparison.OrdinalIgnoreCase);
        Assert.False(previews[1].CanApply);
        Assert.False(previews[2].CanApply);
        Assert.Contains("blank", previews[2].Result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CompatibilityLayer_ConvertsLongInlineValidationWithoutChangingListValues()
    {
        var path = Path.Combine(_tempDirectory, "long-validation.xlsx");
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Import Mapping");
            sheet.Cell("A4").InsertData(new[]
            {
                new[] { "Mapping Profile", "Source Field", "Target Field" },
                new[] { "Profile", "source", "target" }
            });
            workbook.SaveAs(path);
        }

        var values = Enumerable.Range(1, 40).Select(index => $"Long validation option {index:00}").ToArray();
        using (var document = SpreadsheetDocument.Open(path, true))
        {
            var worksheet = document.WorkbookPart!.WorksheetParts.First().Worksheet;
            worksheet.Append(new DataValidations(
                new DataValidation(
                    new Formula1($"\"{string.Join(',', values)}\""))
                {
                    Type = DataValidationValues.List,
                    SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "A1:A2" }
                }) { Count = 1U });
            worksheet.Save();
        }

        var normalized = WorkbookCompatibilityService.NormalizeLongInlineValidations(path);

        Assert.Equal(1, normalized);
        using var result = new XLWorkbook(path);
        var helper = result.Worksheet("CVIS Validation Lists");
        Assert.Equal(XLWorksheetVisibility.VeryHidden, helper.Visibility);
        Assert.Equal(values, helper.Column(1).Cells(1, values.Length).Select(cell => cell.GetString()));
    }

    private string CreateApplyWorkbook()
    {
        var path = Path.Combine(_tempDirectory, $"apply-{Guid.NewGuid():N}.xlsx");
        using var workbook = new XLWorkbook();
        var inventory = workbook.AddWorksheet("Application Inventory");
        inventory.Cell("A1").Value = "Application Inventory";
        inventory.Cell("A4").InsertData(new[]
        {
            new[] { "App ID", "Application / Tool", "Owner", "Status" },
            new[] { "APP-001", "AccountCheck", "Old Owner", "Yellow" }
        });

        var intake = workbook.AddWorksheet("Data Intake");
        intake.Cell("A1").Value = "Data Intake Queue";
        intake.Cell("A4").InsertData(new[] { new[]
        {
            "Intake ID", "Submitted Date", "Submitted By", "Update Type", "Target Table",
            "Record Key Field", "Record Key Value", "Field To Update", "New Value",
            "Effective Date", "Source File", "Source Row", "Review Status",
            "Reviewed By", "Applied Date", "Notes"
        }});

        var log = workbook.AddWorksheet("Change Log");
        log.Cell("A1").Value = "Change Log";
        log.Cell("A4").InsertData(new[] { new[]
        {
            "Change ID", "Timestamp", "Changed By", "Target Table", "Record Key",
            "Field", "Old Value", "New Value", "Source", "Review Status", "Notes"
        }});
        workbook.SaveAs(path);
        return path;
    }

    private static IntakeUpdate OwnerUpdate(string owner) => new()
    {
        TargetTable = "Application Inventory",
        RecordKeyField = "App ID",
        RecordKeyValue = "APP-001",
        FieldToUpdate = "Owner",
        NewValue = owner
    };

    public void Dispose()
    {
        Directory.Delete(_tempDirectory, recursive: true);
    }
}
