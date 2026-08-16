namespace CVIS.WorkbookUpdater.Models;

public sealed class MappingRow
{
    public string MappingProfile { get; set; } = string.Empty;
    public string SourceType { get; set; } = string.Empty;
    public string TargetTable { get; set; } = string.Empty;
    public string SourceField { get; set; } = string.Empty;
    public string TargetField { get; set; } = string.Empty;
    public bool Required { get; set; }
    public bool MatchKey { get; set; }
    public string TransformRule { get; set; } = string.Empty;
    public string DefaultValue { get; set; } = string.Empty;
    public string Notes { get; set; } = string.Empty;
}
