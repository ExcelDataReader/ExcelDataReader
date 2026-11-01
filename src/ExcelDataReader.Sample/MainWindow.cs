using System.Data;
using System.Diagnostics;
using Avalonia.Controls;
using Avalonia.Platform.Storage;

namespace ExcelDataReader.Sample;

internal sealed partial class MainWindow : Window
{
    private static readonly FilePickerFileType SupportedFiles = new("Supported files")
    {
        Patterns = ["*.xls", "*.xlsx", "*.xlsb", "*.xml", "*.csv"]
    };

    private DataSet? _dataSet;

    public MainWindow()
    {
        InitializeComponent();
    }

    private async void SelectFileClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            AllowMultiple = false,
            FileTypeFilter = [SupportedFiles]
        });

        if (files.Count > 0)
        {
            FilePathTextBox.Text = files[0].Path.LocalPath;
        }
    }

    private void ProcessClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try
        {
            var filePath = FilePathTextBox.Text;
            if (string.IsNullOrWhiteSpace(filePath))
            {
                StatusTextBlock.Text = "Select a file first.";
                return;
            }

            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var stopwatch = Stopwatch.StartNew();

            using IExcelDataReader reader = string.Equals(Path.GetExtension(filePath), ".csv", StringComparison.OrdinalIgnoreCase)
                ? ExcelReaderFactory.CreateCsvReader(stream)
                : ExcelReaderFactory.CreateReader(stream, new() { SinglePassMode = SinglePassModeCheckBox.IsChecked == true });

            var openTiming = stopwatch.ElapsedMilliseconds;
            _dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                UseColumnDataType = false,
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = FirstRowNamesCheckBox.IsChecked == true,
                    FillMergedCellsValue = FillMergedCellsCheckBox.IsChecked == true
                }
            });

            stopwatch.Stop();
            StatusTextBlock.Text = $"Elapsed: {stopwatch.ElapsedMilliseconds} ms ({openTiming} ms to open)";

            SheetComboBox.ItemsSource = _dataSet.Tables.Cast<DataTable>().ToList();
            SheetComboBox.SelectedIndex = _dataSet.Tables.Count > 0 ? 0 : -1;
            if (_dataSet.Tables.Count == 0)
            {
                SheetTableView.Columns.Clear();
                SheetTableView.ItemsSource = null;
            }
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"Error: {ex.Message}";
        }
    }

    private void SheetSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (SheetComboBox.SelectedItem is not DataTable table)
        {
            return;
        }

        SheetTableView.Columns.Clear();
        for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
        {
            SheetTableView.Columns.Add(new TableViewColumn
            {
                Header = table.Columns[columnIndex].ColumnName,
                Binding = new Avalonia.Data.Binding($"[{columnIndex}]")
            });
        }

        var rows = new List<SheetGridRow>(table.Rows.Count);
        foreach (DataRow dataRow in table.Rows)
        {
            var values = new object?[table.Columns.Count];
            for (var columnIndex = 0; columnIndex < values.Length; columnIndex++)
            {
                var value = dataRow[columnIndex];
                values[columnIndex] = value == DBNull.Value ? null : value;
            }

            rows.Add(new SheetGridRow(values));
        }

        SheetTableView.ItemsSource = rows;
    }
}
