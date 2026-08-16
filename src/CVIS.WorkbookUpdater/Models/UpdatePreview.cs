namespace CVIS.WorkbookUpdater.Models;

public sealed class UpdatePreview
{
    public required IntakeUpdate Update { get; init; }
    public string CurrentValue { get; init; } = string.Empty;
    public bool TargetFound { get; init; }
    public bool CanApply { get; init; }
    public string Result { get; init; } = string.Empty;
}

public sealed class WorkbookSaveResult
{
    public int UpdateCount { get; init; }
    public string BackupPath { get; init; } = string.Empty;
}
