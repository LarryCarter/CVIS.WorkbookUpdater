using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using CVIS.WorkbookUpdater.Models;
using CVIS.WorkbookUpdater.Services;

namespace CVIS.WorkbookUpdater;

public partial class MainWindow : Window
{
    private readonly WorkbookUpdateService _workbookService = new();
    private readonly ImportService _importService = new();
    private IReadOnlyList<MappingRow> _mappings = Array.Empty<MappingRow>();
    private IReadOnlyList<IntakeUpdate> _pendingImportUpdates = Array.Empty<IntakeUpdate>();

    public MainWindow()
    {
        InitializeComponent();
        SubmittedByTextBox.Text = Environment.UserName;
        EffectiveDatePicker.SelectedDate = DateTime.Today;
        RiskNextDatePicker.SelectedDate = DateTime.Today.AddDays(7);
    }

    private void BrowseWorkbook_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            Title = "Select CVIS recovery workbook"
        };

        if (dialog.ShowDialog() == true)
        {
            WorkbookPathTextBox.Text = dialog.FileName;
        }
    }

    private void BrowseImport_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Supported imports (*.csv;*.yaml;*.yml)|*.csv;*.yaml;*.yml|CSV (*.csv)|*.csv|YAML (*.yaml;*.yml)|*.yaml;*.yml",
            Title = "Select team update import file"
        };

        if (dialog.ShowDialog() == true)
        {
            ImportPathTextBox.Text = dialog.FileName;
        }
    }

    private void LoadMappings_Click(object sender, RoutedEventArgs e)
    {
        TryRun(() =>
        {
            RequireWorkbookPath();
            _mappings = _workbookService.ReadMappings(WorkbookPathTextBox.Text);
            MappingSummaryTextBlock.Text = $"Loaded {_mappings.Count} mapping rows from Import Mapping.";
            Log($"Loaded {_mappings.Count} mapping rows.");
        });
    }

    private void AppendManualUpdate_Click(object sender, RoutedEventArgs e)
    {
        TryRun(() =>
        {
            RequireWorkbookPath();
            var update = new IntakeUpdate
            {
                SubmittedBy = SubmittedByTextBox.Text.Trim(),
                UpdateType = ComboValue(UpdateTypeComboBox, "Manual Update"),
                TargetTable = ComboValue(TargetTableComboBox, "Application Inventory"),
                RecordKeyField = RecordKeyFieldTextBox.Text.Trim(),
                RecordKeyValue = RecordKeyValueTextBox.Text.Trim(),
                FieldToUpdate = FieldToUpdateTextBox.Text.Trim(),
                NewValue = NewValueTextBox.Text.Trim(),
                EffectiveDate = EffectiveDatePicker.SelectedDate,
                Notes = NotesTextBox.Text.Trim()
            };

            ValidateUpdate(update);
            var count = _workbookService.AppendIntakeUpdates(WorkbookPathTextBox.Text, [update], DryRunCheckBox.IsChecked == true);
            Log($"{count} manual update row appended to Data Intake.");
        });
    }

    private void ApplyManualUpdate_Click(object sender, RoutedEventArgs e)
    {
        TryRun(() =>
        {
            RequireWorkbookPath();
            var update = BuildManualUpdate();
            ValidateUpdate(update);
            PreviewConfirmAndApply([update], "manual update");
        });
    }

    private IntakeUpdate BuildManualUpdate() => new()
    {
        SubmittedBy = SubmittedByTextBox.Text.Trim(),
        UpdateType = ComboValue(UpdateTypeComboBox, "Manual Update"),
        TargetTable = ComboValue(TargetTableComboBox, "Application Inventory"),
        RecordKeyField = RecordKeyFieldTextBox.Text.Trim(),
        RecordKeyValue = RecordKeyValueTextBox.Text.Trim(),
        FieldToUpdate = FieldToUpdateTextBox.Text.Trim(),
        NewValue = NewValueTextBox.Text.Trim(),
        EffectiveDate = EffectiveDatePicker.SelectedDate,
        Notes = NotesTextBox.Text.Trim()
    };

    private void ImportUpdates_Click(object sender, RoutedEventArgs e)
    {
        TryRun(() =>
        {
            RequireWorkbookPath();
            if (!File.Exists(ImportPathTextBox.Text))
            {
                throw new InvalidOperationException("Select a CSV or YAML import file.");
            }

            var updates = BuildImportUpdates();
            var count = _workbookService.AppendIntakeUpdates(WorkbookPathTextBox.Text, updates, dryRun: false);
            Log($"{count} imported update rows appended to Data Intake.");
        });
    }

    private void PreviewImport_Click(object sender, RoutedEventArgs e)
    {
        TryRun(() =>
        {
            RequireWorkbookPath();
            _pendingImportUpdates = BuildImportUpdates();
            var previews = _workbookService.PreviewUpdates(WorkbookPathTextBox.Text, _pendingImportUpdates);
            ImportPreviewGrid.ItemsSource = previews;
            var ready = previews.Count(preview => preview.CanApply);
            MappingSummaryTextBlock.Text = $"{previews.Count} proposed updates: {ready} ready, {previews.Count - ready} need attention.";
            Log($"Previewed {previews.Count} imported updates; {ready} are ready to apply.");
        });
    }

    private void ApplyImportUpdates_Click(object sender, RoutedEventArgs e)
    {
        TryRun(() =>
        {
            RequireWorkbookPath();
            if (_pendingImportUpdates.Count == 0)
            {
                _pendingImportUpdates = BuildImportUpdates();
            }

            PreviewConfirmAndApply(_pendingImportUpdates, "imported updates");
            ImportPreviewGrid.ItemsSource = _workbookService.PreviewUpdates(WorkbookPathTextBox.Text, _pendingImportUpdates);
        });
    }

    private IReadOnlyList<IntakeUpdate> BuildImportUpdates()
    {
        if (!File.Exists(ImportPathTextBox.Text))
        {
            throw new InvalidOperationException("Select a CSV or YAML import file.");
        }

        if (_mappings.Count == 0)
        {
            _mappings = _workbookService.ReadMappings(WorkbookPathTextBox.Text);
        }

        return _importService.BuildUpdatesFromFile(
            ImportPathTextBox.Text,
            SubmittedByTextBox.Text.Trim(),
            _mappings,
            MappingProfileTextBox.Text.Trim());
    }

    private void ApplyRiskUpdate_Click(object sender, RoutedEventArgs e)
    {
        TryRun(() =>
        {
            RequireWorkbookPath();
            var item = RiskItemTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(item))
            {
                throw new InvalidOperationException("Risk / Item is required.");
            }

            var status = ComboValue(RiskStatusComboBox, "Yellow");
            if (!status.Equals("Green", StringComparison.OrdinalIgnoreCase)
                && (string.IsNullOrWhiteSpace(RiskReasonTextBox.Text) || string.IsNullOrWhiteSpace(RiskActionTextBox.Text)))
            {
                throw new InvalidOperationException("Red and Yellow updates require both a reason/blocker and a recovery action.");
            }

            var updates = new List<IntakeUpdate>
            {
                QuickUpdate("At-Risk Register", "Item", item, "Current RYG", status, "RYG / Risk Update", RiskReasonTextBox.Text)
            };
            AddIfPopulated(updates, "At-Risk Register", "Item", item, "Reason", RiskReasonTextBox.Text, "Blocker Update");
            AddIfPopulated(updates, "At-Risk Register", "Item", item, "Action Plan", RiskActionTextBox.Text, "Recovery Update");
            AddIfPopulated(updates, "At-Risk Register", "Item", item, "Leadership Ask", RiskLeadershipAskTextBox.Text, "Leadership Ask");
            if (RiskNextDatePicker.SelectedDate is DateTime nextDate)
            {
                updates.Add(QuickUpdate("At-Risk Register", "Item", item, "Next Action Date",
                    nextDate.ToString("yyyy-MM-dd"), "Date Update", RiskReasonTextBox.Text));
            }

            PreviewConfirmAndApply(updates, "risk update");
        });
    }

    private void ApplyWorkUpdate_Click(object sender, RoutedEventArgs e)
    {
        TryRun(() =>
        {
            RequireWorkbookPath();
            var target = ComboValue(WorkTargetComboBox, "Action Plan");
            var keyField = WorkKeyFieldTextBox.Text.Trim();
            var keyValue = WorkIdTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(keyField) || string.IsNullOrWhiteSpace(keyValue))
            {
                throw new InvalidOperationException("Key Field and Ticket / Work ID are required.");
            }

            var updates = new List<IntakeUpdate>();
            AddIfPopulated(updates, target, keyField, keyValue, WorkStatusFieldTextBox.Text,
                ComboValue(WorkStatusComboBox, "In Progress"), "Status Update", WorkNotesTextBox.Text);
            AddIfPopulated(updates, target, keyField, keyValue, WorkNextActionFieldTextBox.Text,
                WorkNextActionTextBox.Text, "Work Update", WorkNotesTextBox.Text);
            if (updates.Count == 0)
            {
                throw new InvalidOperationException("Enter a status or next action to update.");
            }

            PreviewConfirmAndApply(updates, "ticket/work update");
        });
    }

    private void PreviewConfirmAndApply(IReadOnlyList<IntakeUpdate> updates, string description)
    {
        var previews = _workbookService.PreviewUpdates(WorkbookPathTextBox.Text, updates);
        var blocked = previews.Where(preview => !preview.CanApply).ToList();
        if (blocked.Count > 0)
        {
            var details = string.Join(Environment.NewLine, blocked.Take(8).Select(preview =>
                $"{preview.Update.TargetTable} / {preview.Update.RecordKeyValue} / {preview.Update.FieldToUpdate}: {preview.Result}"));
            throw new InvalidOperationException($"Cannot apply {description}:{Environment.NewLine}{details}");
        }

        var changes = previews.Count(preview => !preview.Result.StartsWith("No change", StringComparison.OrdinalIgnoreCase));
        var confirmation = MessageBox.Show(
            $"Apply {updates.Count} {description} ({changes} value changes)? A verified backup will be created first.",
            "Confirm workbook update",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirmation != MessageBoxResult.Yes)
        {
            Log($"Cancelled {description}.");
            return;
        }

        var result = _workbookService.ApplyUpdates(
            WorkbookPathTextBox.Text,
            updates,
            string.IsNullOrWhiteSpace(SubmittedByTextBox.Text) ? Environment.UserName : SubmittedByTextBox.Text.Trim());
        Log($"Applied {result.UpdateCount} {description}. Backup: {result.BackupPath}");
    }

    private IntakeUpdate QuickUpdate(
        string target,
        string keyField,
        string keyValue,
        string field,
        string value,
        string updateType,
        string notes) => new()
    {
        SubmittedBy = string.IsNullOrWhiteSpace(SubmittedByTextBox.Text) ? Environment.UserName : SubmittedByTextBox.Text.Trim(),
        UpdateType = updateType,
        TargetTable = target.Trim(),
        RecordKeyField = keyField.Trim(),
        RecordKeyValue = keyValue.Trim(),
        FieldToUpdate = field.Trim(),
        NewValue = value.Trim(),
        EffectiveDate = DateTime.Today,
        Notes = notes.Trim()
    };

    private void AddIfPopulated(
        ICollection<IntakeUpdate> updates,
        string target,
        string keyField,
        string keyValue,
        string field,
        string value,
        string updateType,
        string notes = "")
    {
        if (!string.IsNullOrWhiteSpace(field) && !string.IsNullOrWhiteSpace(value))
        {
            updates.Add(QuickUpdate(target, keyField, keyValue, field, value, updateType, notes));
        }
    }

    private void RequireWorkbookPath()
    {
        var value = WorkbookPathTextBox.Text.Trim();
        var isHttpUrl = Uri.TryCreate(value, UriKind.Absolute, out var uri)
            && (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp);
        if (string.IsNullOrWhiteSpace(value) || (!isHttpUrl && !File.Exists(value)))
        {
            throw new InvalidOperationException("Enter a direct workbook URL or select the CVIS recovery workbook first.");
        }
    }

    private static void ValidateUpdate(IntakeUpdate update)
    {
        if (string.IsNullOrWhiteSpace(update.RecordKeyValue))
        {
            throw new InvalidOperationException("Record Key Value is required.");
        }

        if (string.IsNullOrWhiteSpace(update.FieldToUpdate))
        {
            throw new InvalidOperationException("Field To Update is required.");
        }
    }

    private static string ComboValue(ComboBox comboBox, string fallback)
    {
        var value = comboBox.SelectedItem is ComboBoxItem item
            ? item.Content?.ToString()
            : comboBox.Text;

        return string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
    }

    private void TryRun(Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            Log($"ERROR: {ex.Message}");
            MessageBox.Show(ex.Message, "CVIS Workbook Updater", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void Log(string message)
    {
        LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
        LogTextBox.ScrollToEnd();
    }
}
