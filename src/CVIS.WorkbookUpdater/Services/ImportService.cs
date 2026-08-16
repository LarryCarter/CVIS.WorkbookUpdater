using System.Globalization;
using System.IO;
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

        var keyMappings = profileMappings.Where(m => m.MatchKey).ToList();
        if (keyMappings.Count == 0)
        {
            throw new InvalidOperationException($"Mapping profile '{preferredProfile}' needs at least one row marked Match Key = Yes.");
        }

        var updates = new List<IntakeUpdate>();
        for (var recordIndex = 0; recordIndex < records.Count; recordIndex++)
        {
            var record = records[recordIndex];
            var selectedKey = keyMappings
                .Select(mapping => new
                {
                    Mapping = mapping,
                    Value = ApplyTransform(ReadValue(record, mapping.SourceField), mapping.TransformRule)
                })
                .FirstOrDefault(candidate => !string.IsNullOrWhiteSpace(candidate.Value));

            if (selectedKey is null)
            {
                var expectedKeys = string.Join(", ", keyMappings.Select(mapping => mapping.SourceField));
                throw new InvalidOperationException(
                    $"Source row {recordIndex + 2} has no usable match key. Expected one of: {expectedKeys}.");
            }

            foreach (var mapping in profileMappings.Where(m => !m.MatchKey))
            {
                var newValue = ReadValue(record, mapping.SourceField);
                if (string.IsNullOrWhiteSpace(newValue))
                {
                    newValue = mapping.DefaultValue;
                }

                if (string.IsNullOrWhiteSpace(newValue) && mapping.Required)
                {
                    throw new InvalidOperationException(
                        $"Source row {recordIndex + 2} is missing required field '{mapping.SourceField}'.");
                }

                if (string.IsNullOrWhiteSpace(newValue))
                {
                    continue;
                }

                updates.Add(new IntakeUpdate
                {
                    SubmittedBy = submittedBy,
                    UpdateType = "Bulk Import",
                    TargetTable = mapping.TargetTable,
                    RecordKeyField = selectedKey.Mapping.TargetField,
                    RecordKeyValue = selectedKey.Value,
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
        var trimmed = value.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return string.Empty;
        }

        return transformRule.Trim().ToLowerInvariant() switch
        {
            "uppercase" => trimmed.ToUpperInvariant(),
            "lowercase" => trimmed.ToLowerInvariant(),
            "number" => NormalizeNumber(trimmed),
            "date" => NormalizeDate(trimmed),
            "normalize ryg" => NormalizeRyg(trimmed),
            "normalize status" => NormalizeStatus(trimmed),
            "normalize list" or "append note" or "trim" or "" => CollapseWhitespace(trimmed),
            _ => throw new InvalidOperationException($"Unsupported transform rule '{transformRule}'.")
        };
    }

    private static string NormalizeNumber(string value)
    {
        return decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var number)
            ? number.ToString(CultureInfo.InvariantCulture)
            : throw new InvalidOperationException($"'{value}' is not a valid number.");
    }

    private static string NormalizeDate(string value)
    {
        return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var date)
            ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : throw new InvalidOperationException($"'{value}' is not a valid date.");
    }

    private static string NormalizeRyg(string value)
    {
        return value.Trim().ToLowerInvariant() switch
        {
            "r" or "red" => "Red",
            "y" or "yellow" or "amber" => "Yellow",
            "g" or "green" => "Green",
            _ => throw new InvalidOperationException($"'{value}' is not a valid RYG status.")
        };
    }

    private static string NormalizeStatus(string value)
    {
        return value.Trim().ToLowerInvariant() switch
        {
            "todo" or "to do" or "not started" => "To do",
            "in progress" or "active" => "In Progress",
            "on hold" or "waiting" or "waiting external" => "On Hold",
            "blocked" => "Blocked",
            "done" or "complete" or "completed" => "Done",
            _ => throw new InvalidOperationException($"'{value}' is not a recognized status.")
        };
    }

    private static string CollapseWhitespace(string value)
    {
        return string.Join(' ', value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }
}
