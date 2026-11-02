using System.Data;
using System.Diagnostics;
using Avalonia.Collections;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Platform.Storage;

namespace ExcelDataReader.Sample;

internal sealed partial class MainWindow : Window
{
    private static readonly FilePickerFileType SupportedFiles = new("Supported files")
    {
        Patterns = ["*.xls", "*.xlsx", "*.xlsb", "*.xml", "*.csv"]
    };

    private DataSet? _dataSet;
    private string? _selectedPath;
    private bool _isProcessing;

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
            _selectedPath = files[0].Path.LocalPath;
            FilePathTextBox.Text = _selectedPath;
        }
    }

    private async void ProcessClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (_isProcessing)
        {
            return;
        }

        var filePath = _selectedPath;
        if (string.IsNullOrWhiteSpace(filePath))
        {
            StatusTextBlock.Text = "Choose a file first.";
            return;
        }

        var useHeaderRow = FirstRowNamesCheckBox.IsChecked == true;
        var fillMergedCellsValue = FillMergedCellsCheckBox.IsChecked == true;
        var singlePassMode = SinglePassModeCheckBox.IsChecked == true;
        _isProcessing = true;
        LoadControls.IsEnabled = false;

        try
        {
            SheetTableView.ItemsSource = null;
            SheetTableView.Columns = new AvaloniaList<TableViewColumn>();
            SheetComboBox.SelectedIndex = -1;
            SheetComboBox.ItemsSource = null;
            _dataSet?.Dispose();
            _dataSet = null;
            StatusTextBlock.Text = "Loading...";

            var (dataSet, elapsed, openTiming) = await Task.Run(() =>
            {
                using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var stopwatch = Stopwatch.StartNew();

                using IExcelDataReader reader = string.Equals(Path.GetExtension(filePath), ".csv", StringComparison.OrdinalIgnoreCase)
                    ? ExcelReaderFactory.CreateCsvReader(stream)
                    : ExcelReaderFactory.CreateReader(stream, new() { SinglePassMode = singlePassMode });

                var open = stopwatch.ElapsedMilliseconds;
                var result = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    UseColumnDataType = false,
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = useHeaderRow,
                        FillMergedCellsValue = fillMergedCellsValue
                    }
                });

                return (result, stopwatch.ElapsedMilliseconds, open);
            });

            _dataSet = dataSet;
            StatusTextBlock.Text = $"Elapsed: {elapsed} ms ({openTiming} ms to open)";

            SheetComboBox.ItemsSource = dataSet.Tables.Cast<DataTable>().ToList();
            SheetComboBox.SelectedIndex = dataSet.Tables.Count > 0 ? 0 : -1;
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"Error: {ex.Message}";
        }
        finally
        {
            _isProcessing = false;
            LoadControls.IsEnabled = true;
        }
    }

    private void SheetSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (SheetComboBox.SelectedItem is not DataTable table)
        {
            return;
        }

        SheetTableView.ItemsSource = null;
        var columns = new AvaloniaList<TableViewColumn>();
        for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
        {
            var index = columnIndex;
            columns.Add(new TableViewColumn
            {
                Header = table.Columns[columnIndex].ColumnName,
                Width = new GridLength(140),
                CellTemplate = new FuncDataTemplate<DataRow>((row, _) => new TextBlock
                {
                    Text = row[index] is DBNull ? null : row[index].ToString()
                })
            });
        }

        SheetTableView.Columns = columns;
        SheetTableView.ItemsSource = new SheetRows(table.Rows);
    }
}
