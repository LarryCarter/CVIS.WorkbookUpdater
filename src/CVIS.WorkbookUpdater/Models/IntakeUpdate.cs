namespace CVIS.WorkbookUpdater.Models;

public sealed class IntakeUpdate
{
    public string IntakeId { get; set; } = string.Empty;
    public DateTime SubmittedDate { get; set; } = DateTime.Now;
    public string SubmittedBy { get; set; } = Environment.UserName;
    public string UpdateType { get; set; } = "Manual Update";
    public string TargetTable { get; set; } = "Application Inventory";
    public string RecordKeyField { get; set; } = "Application / Component";
    public string RecordKeyValue { get; set; } = string.Empty;
    public string FieldToUpdate { get; set; } = string.Empty;
    public string NewValue { get; set; } = string.Empty;
    public DateTime? EffectiveDate { get; set; } = DateTime.Today;
    public string SourceFile { get; set; } = string.Empty;
    public string SourceRow { get; set; } = string.Empty;
    public string ReviewStatus { get; set; } = "Needs Review";
    public string ReviewedBy { get; set; } = string.Empty;
    public DateTime? AppliedDate { get; set; }
    public string Notes { get; set; } = string.Empty;
}
