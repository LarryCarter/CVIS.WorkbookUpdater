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

    public MainWindow()
    {
        InitializeComponent();
        SubmittedByTextBox.Text = Environment.UserName;
        EffectiveDatePicker.SelectedDate = DateTime.Today;
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

    private void ImportUpdates_Click(object sender, RoutedEventArgs e)
    {
        TryRun(() =>
        {
            RequireWorkbookPath();
            if (!File.Exists(ImportPathTextBox.Text))
            {
                throw new InvalidOperationException("Select a CSV or YAML import file.");
            }

            if (_mappings.Count == 0)
            {
                _mappings = _workbookService.ReadMappings(WorkbookPathTextBox.Text);
            }

            var updates = _importService.BuildUpdatesFromFile(
                ImportPathTextBox.Text,
                SubmittedByTextBox.Text.Trim(),
                _mappings,
                MappingProfileTextBox.Text.Trim());

            var count = _workbookService.AppendIntakeUpdates(WorkbookPathTextBox.Text, updates, DryRunCheckBox.IsChecked == true);
            Log($"{count} imported update rows appended to Data Intake.");
        });
    }

    private void RequireWorkbookPath()
    {
        if (string.IsNullOrWhiteSpace(WorkbookPathTextBox.Text) || !File.Exists(WorkbookPathTextBox.Text))
        {
            throw new InvalidOperationException("Select the CVIS recovery workbook first.");
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
