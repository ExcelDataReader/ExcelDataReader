using System.Xml;
using ExcelDataReader.Core.NumberFormat;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat;

internal sealed class XlsxWorksheet : IWorksheet
{
    public XlsxWorksheet(ZipWorker document, XlsxWorkbook workbook, SheetRecord refSheet, bool singlePassMode = false)
    {
        Document = document;
        Workbook = workbook;

        Name = refSheet.Name;
        VisibleState = refSheet.VisibleState;
        Path = refSheet.Path;
        DefaultRowHeight = 15;

        if (string.IsNullOrEmpty(Path))
            return;

        using var sheetStream = Document.GetWorksheetReader(Path, !singlePassMode);
        
        if (sheetStream == null)
            return;

        int rowIndexMaximum = int.MinValue;
        int columnIndexMaximum = int.MinValue;

        List<Column> columnWidths = [];
        List<CellRange> cellRanges = [];

        bool stop = false;

        while (!stop && sheetStream.Read() is { } record)
        {
            switch (record)
            {
                case SheetDimRecord dimRecord:
                    Dimension = dimRecord.Range;
                    break;
                case SheetDataBeginRecord _ when singlePassMode:
                    // In single-pass mode stop before reading cells; ReadRows() will be the only pass
                    stop = true;
                    break;
                case ScanResultRecord scan:
                    // Multi-pass: sheetData was scanned without allocating per-cell records
                    rowIndexMaximum = Math.Max(rowIndexMaximum, scan.MaxRowIndex);
                    columnIndexMaximum = Math.Max(columnIndexMaximum, scan.MaxColumnIndex);
                    break;
                case ColumnRecord column:
                    columnWidths.Add(column.Column);
                    break;
                case SheetFormatPrRecord sheetFormatProperties:
                    if (sheetFormatProperties.DefaultRowHeight != null)
                        DefaultRowHeight = sheetFormatProperties.DefaultRowHeight.Value;
                    break;
                case SheetPrRecord sheetProperties:
                    CodeName = sheetProperties.CodeName;
                    break;
                case MergeCellRecord mergeCell:
                    cellRanges.Add(mergeCell.Range);
                    break;
                case HeaderFooterRecord headerFooter:
                    HeaderFooter = headerFooter.HeaderFooter;
                    break;
            }
        }

        ColumnWidths = columnWidths;
        MergeCells = [.. cellRanges];

        if (singlePassMode)
        {
            // In single-pass mode FieldCount comes from the dimension hint only (may be 0 if absent)
            FieldCount = columnIndexMaximum != int.MinValue ? columnIndexMaximum + 1 : 0;
        }
        else if (rowIndexMaximum != int.MinValue && columnIndexMaximum != int.MinValue)
        {
            FieldCount = columnIndexMaximum + 1;
            RowCount = rowIndexMaximum + 1;
        }
    }

    public int FieldCount { get; }

    public int RowCount { get; }

    public CellRange Dimension { get; private set; }

    public string Name { get; }

    public string CodeName { get; }

    public string VisibleState { get; }

    public HeaderFooter HeaderFooter { get; }

    public CellRange[] MergeCells { get; }

    public List<Column> ColumnWidths { get; }

    private string Path { get; set; }

    private double DefaultRowHeight { get; }

    private ZipWorker Document { get; }

    private XlsxWorkbook Workbook { get; }

    public IEnumerable<Row> ReadRows()
    {
        if (string.IsNullOrEmpty(Path))
            yield break;

        using RecordReader sheetStream = Document.GetWorksheetReader(Path, false);
        if (sheetStream == null)
            yield break;

        var rowIndex = 0;
        List<Cell> cells = [];
        bool foundRowOrCell = false;
        double height = 0;

        bool inSheetData = false;
        while (sheetStream.Read() is { } record)
        {
            switch (record)
            {
                case SheetDataBeginRecord _:
                    inSheetData = true;
                    break;
                case SheetDataEndRecord _:
                    inSheetData = false;
                    break;
                case RowHeaderRecord row when inSheetData:
                    foundRowOrCell = true;

                    int currentRowIndex = row.RowIndex;
                    if (rowIndex != currentRowIndex)
                    {
                        yield return new Row(rowIndex++, height, cells);
                        cells.Clear();
                    }

                    for (; rowIndex < currentRowIndex; rowIndex++)
                    {
                        yield return new Row(rowIndex, DefaultRowHeight, cells);
                    }

                    height = row.Hidden ? 0 : row.Height ?? DefaultRowHeight;

                    break;
                case CellRecord cell when inSheetData:
                    // TODO What if we get a cell without a row?
                    var extendedFormat = Workbook.GetEffectiveCellStyle(cell.XfIndex, 0);
                    cells.Add(new Cell(cell.ColumnIndex, ConvertCellValue(cell.Value, extendedFormat.NumberFormatIndex), extendedFormat, cell.Error));
                    foundRowOrCell = true;
                    break;
            }
        }

        if (foundRowOrCell)
            yield return new Row(rowIndex, height, cells);
    }

    private static bool TryParseToTimeSpan(string s, out TimeSpan result)
    {
        var isIsoFormat = Helpers.StringStartsWith(s, 'P');

        if (!isIsoFormat)
        {
            return TimeSpan.TryParse(s, out result);
        }

        try
        {
            result = XmlConvert.ToTimeSpan(s);
            return true;
        }
        catch (FormatException)
        {
            result = TimeSpan.Zero;
            return false;
        }
    }

    private object ConvertCellValue(object value, int numberFormatIndex)
    {
        switch (value)
        {
            case int sstIndex:
                if (sstIndex >= 0 && sstIndex < Workbook.SST.Count)
                {
                    return Helpers.ConvertEscapeChars(Workbook.SST[sstIndex]);
                }

                return null;

            case double number:
                var format = Workbook.GetNumberFormatString(numberFormatIndex, null);
                if (format != null)
                {
                    if (format.IsDateTimeFormat)
                        return Helpers.ConvertFromOATime(number, Workbook.IsDate1904);
                    if (format.IsTimeSpanFormat)
                        return TimeSpan.FromDays(number);
                }

                return number;

            case DateTime date:
                return date;

            case string s:
                NumberFormatString numberFormat = Workbook.GetNumberFormatString(numberFormatIndex, null);
                if (numberFormat.IsTimeSpanFormat && TryParseToTimeSpan(s, out var timeSpan))
                {
                    return timeSpan;
                }

                if (numberFormat.IsDateTimeFormat && DateTime.TryParse(s, out DateTime dateTime))
                {
                    return dateTime;
                }

                return s;

            default:
                return value;
        }
    }
}
