using System.Globalization;
using CsvHelper;
using CVIS.WorkbookUpdater.Models;
using YamlDotNet.Serialization;

namespace CVIS.WorkbookUpdater.Services;

public sealed class ImportService
{
    public IReadOnlyList<IntakeUpdate> BuildUpdatesFromFile(
        string sourcePath,
        string submittedBy,
        IReadOnlyList<MappingRow> mappings,
        string preferredProfile)
    {
        var extension = Path.GetExtension(sourcePath).ToLowerInvariant();
        var records = extension switch
        {
            ".csv" => ReadCsv(sourcePath),
            ".yaml" or ".yml" => ReadYaml(sourcePath),
            _ => throw new InvalidOperationException("Only CSV, YAML, and YML files are supported.")
        };

        var profileMappings = mappings
            .Where(m => m.MappingProfile.Equals(preferredProfile, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (profileMappings.Count == 0)
        {
            throw new InvalidOperationException($"No mappings found for profile '{preferredProfile}'. Add rows to the Import Mapping sheet first.");
        }

        var keyMapping = profileMappings.FirstOrDefault(m => m.MatchKey)
            ?? throw new InvalidOperationException($"Mapping profile '{preferredProfile}' needs one row marked Match Key = Yes.");

        var updates = new List<IntakeUpdate>();
        for (var recordIndex = 0; recordIndex < records.Count; recordIndex++)
        {
            var record = records[recordIndex];
            var recordKeyValue = ReadValue(record, keyMapping.SourceField);
            if (string.IsNullOrWhiteSpace(recordKeyValue))
            {
                continue;
            }

            foreach (var mapping in profileMappings.Where(m => !m.MatchKey))
            {
                var newValue = ReadValue(record, mapping.SourceField);
                if (string.IsNullOrWhiteSpace(newValue))
                {
                    newValue = mapping.DefaultValue;
                }

                if (string.IsNullOrWhiteSpace(newValue) && !mapping.Required)
                {
                    continue;
                }

                updates.Add(new IntakeUpdate
                {
                    SubmittedBy = submittedBy,
                    UpdateType = "Bulk Import",
                    TargetTable = mapping.TargetTable,
                    RecordKeyField = keyMapping.TargetField,
                    RecordKeyValue = recordKeyValue,
                    FieldToUpdate = mapping.TargetField,
                    NewValue = ApplyTransform(newValue, mapping.TransformRule),
                    SourceFile = Path.GetFileName(sourcePath),
                    SourceRow = (recordIndex + 2).ToString(CultureInfo.InvariantCulture),
                    Notes = mapping.Notes
                });
            }
        }

        return updates;
    }

    private static IReadOnlyList<Dictionary<string, string>> ReadCsv(string path)
    {
        using var reader = new StreamReader(path);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        var rows = new List<Dictionary<string, string>>();
        foreach (var row in csv.GetRecords<dynamic>())
        {
            var dictionary = (IDictionary<string, object?>)row;
            rows.Add(dictionary.ToDictionary(
                kvp => kvp.Key,
                kvp => kvp.Value?.ToString() ?? string.Empty,
                StringComparer.OrdinalIgnoreCase));
        }

        return rows;
    }

    private static IReadOnlyList<Dictionary<string, string>> ReadYaml(string path)
    {
        var deserializer = new DeserializerBuilder().Build();
        using var reader = new StreamReader(path);
        var yamlObject = deserializer.Deserialize<object>(reader);

        var recordsObject = yamlObject is Dictionary<object, object> root && root.TryGetValue("records", out var records)
            ? records
            : yamlObject;

        if (recordsObject is not IEnumerable<object> sequence)
        {
            throw new InvalidOperationException("YAML must be a list of records or contain a top-level 'records' list.");
        }

        return sequence
            .OfType<Dictionary<object, object>>()
            .Select(row => row.ToDictionary(
                kvp => kvp.Key.ToString() ?? string.Empty,
                kvp => kvp.Value?.ToString() ?? string.Empty,
                StringComparer.OrdinalIgnoreCase))
            .ToList();
    }

    private static string ReadValue(Dictionary<string, string> record, string sourceField)
    {
        return record.TryGetValue(sourceField, out var value) ? value.Trim() : string.Empty;
    }

    private static string ApplyTransform(string value, string transformRule)
    {
        return transformRule.Trim().ToLowerInvariant() switch
        {
            "uppercase" => value.ToUpperInvariant(),
            "lowercase" => value.ToLowerInvariant(),
            "trim" or "" => value.Trim(),
            _ => value.Trim()
        };
    }
}
