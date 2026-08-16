using CVIS.WorkbookUpdater.Models;
using CVIS.WorkbookUpdater.Services;
using Xunit;

namespace CVIS.WorkbookUpdater.Tests;

public sealed class ImportServiceTests : IDisposable
{
    private readonly string _tempDirectory = Path.Combine(
        Path.GetTempPath(),
        $"cvis-import-tests-{Guid.NewGuid():N}");

    public ImportServiceTests()
    {
        Directory.CreateDirectory(_tempDirectory);
    }

    [Fact]
    public void CsvAndYaml_ProduceEquivalentNormalizedUpdates()
    {
        var csvPath = Path.Combine(_tempDirectory, "intake.csv");
        var yamlPath = Path.Combine(_tempDirectory, "intake.yaml");
        File.WriteAllText(csvPath,
            "application_name,app_id,business_function,owner,phase\n" +
            "  AccountCheck  ,app-001,Validate account state,Larry Carter,Phase 1\n");
        File.WriteAllText(yamlPath,
            "records:\n" +
            "  - application_name: '  AccountCheck  '\n" +
            "    app_id: app-001\n" +
            "    business_function: Validate account state\n" +
            "    owner: Larry Carter\n" +
            "    phase: Phase 1\n");

        var service = new ImportService();
        var mappings = ApplicationMappings();

        var csv = service.BuildUpdatesFromFile(csvPath, "Tester", mappings, "ApplicationOwnerIntake");
        var yaml = service.BuildUpdatesFromFile(yamlPath, "Tester", mappings, "ApplicationOwnerIntake");

        Assert.Equal(Project(csv), Project(yaml));
        Assert.All(csv, update => Assert.Equal("APP-001", update.RecordKeyValue));
        Assert.All(csv, update => Assert.Equal("App ID", update.RecordKeyField));
        Assert.Contains(csv, update => update.FieldToUpdate == "Application / Tool" && update.NewValue == "AccountCheck");
    }

    [Fact]
    public void Import_UsesSecondaryMatchKeyWhenPrimaryIsBlank()
    {
        var csvPath = Path.Combine(_tempDirectory, "secondary-key.csv");
        File.WriteAllText(csvPath,
            "application_name,app_id,source_row,business_function,owner,phase\n" +
            "AccountCheck,,42,Validate account state,Larry Carter,Phase 1\n");

        var updates = new ImportService().BuildUpdatesFromFile(
            csvPath, "Tester", ApplicationMappings(), "ApplicationOwnerIntake");

        Assert.All(updates, update => Assert.Equal("Source Row", update.RecordKeyField));
        Assert.All(updates, update => Assert.Equal("42", update.RecordKeyValue));
    }

    [Fact]
    public void Import_ReportsMissingRequiredFieldWithSourceRow()
    {
        var csvPath = Path.Combine(_tempDirectory, "missing-required.csv");
        File.WriteAllText(csvPath,
            "application_name,app_id,business_function,owner,phase\n" +
            "AccountCheck,APP-001,,Larry Carter,Phase 1\n");

        var error = Assert.Throws<InvalidOperationException>(() =>
            new ImportService().BuildUpdatesFromFile(
                csvPath, "Tester", ApplicationMappings(), "ApplicationOwnerIntake"));

        Assert.Contains("Source row 2", error.Message);
        Assert.Contains("business_function", error.Message);
    }

    [Fact]
    public void Import_NormalizesRygStatusAndDate()
    {
        var csvPath = Path.Combine(_tempDirectory, "risk.csv");
        File.WriteAllText(csvPath,
            "item,status,next_action_date\nContainer onboarding,amber,8/31/2026\n");
        var mappings = new[]
        {
            Mapping("AtRiskUpdate", "item", "Item", required: true, matchKey: true),
            Mapping("AtRiskUpdate", "status", "Current RYG", required: true, transform: "Normalize RYG"),
            Mapping("AtRiskUpdate", "next_action_date", "Next Action Date", required: true, transform: "Date")
        };

        var updates = new ImportService().BuildUpdatesFromFile(csvPath, "Tester", mappings, "AtRiskUpdate");

        Assert.Contains(updates, update => update.FieldToUpdate == "Current RYG" && update.NewValue == "Yellow");
        Assert.Contains(updates, update => update.FieldToUpdate == "Next Action Date" && update.NewValue == "2026-08-31");
    }

    private static IReadOnlyList<MappingRow> ApplicationMappings() =>
    [
        Mapping("ApplicationOwnerIntake", "app_id", "App ID", matchKey: true, transform: "Uppercase"),
        Mapping("ApplicationOwnerIntake", "source_row", "Source Row", matchKey: true, transform: "Number"),
        Mapping("ApplicationOwnerIntake", "application_name", "Application / Tool", required: true),
        Mapping("ApplicationOwnerIntake", "business_function", "Business Function", required: true),
        Mapping("ApplicationOwnerIntake", "owner", "Owner", required: true),
        Mapping("ApplicationOwnerIntake", "phase", "Phase", required: true, transform: "Normalize List")
    ];

    private static MappingRow Mapping(
        string profile,
        string source,
        string target,
        bool required = false,
        bool matchKey = false,
        string transform = "Trim") => new()
    {
        MappingProfile = profile,
        SourceType = "CSV/YAML",
        TargetTable = profile == "AtRiskUpdate" ? "At-Risk Register" : "Application Inventory",
        SourceField = source,
        TargetField = target,
        Required = required,
        MatchKey = matchKey,
        TransformRule = transform
    };

    private static string[] Project(IReadOnlyList<IntakeUpdate> updates) => updates
        .Select(update => $"{update.RecordKeyField}|{update.RecordKeyValue}|{update.FieldToUpdate}|{update.NewValue}")
        .ToArray();

    public void Dispose()
    {
        Directory.Delete(_tempDirectory, recursive: true);
    }
}
