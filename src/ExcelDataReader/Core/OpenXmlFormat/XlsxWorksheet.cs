using System.Xml;
using ExcelDataReader.Core.NumberFormat;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat;

internal sealed class XlsxWorksheet : IWorksheet
{
    public XlsxWorksheet(ZipWorker document, XlsxWorkbook workbook, SheetRecord refSheet)
    {
        Document = document;
        Workbook = workbook;

        Name = refSheet.Name;
        VisibleState = refSheet.VisibleState;
        Path = refSheet.Path;
        DefaultRowHeight = 15;

        if (string.IsNullOrEmpty(Path))
            return;

        using var sheetStream = Document.GetWorksheetReader(Path, true);
        
        if (sheetStream == null)
            return;

        int rowIndexMaximum = int.MinValue;
        int columnIndexMaximum = int.MinValue;

        List<Column> columnWidths = [];
        List<CellRange> cellRanges = [];

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
                    rowIndexMaximum = Math.Max(rowIndexMaximum, row.RowIndex);
                    break;
                case CellRecord cell when inSheetData:
                    if (cell.Value != null || cell.Error != null)
                        columnIndexMaximum = Math.Max(columnIndexMaximum, cell.ColumnIndex);
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
                case SheetDimRecord dimRecord:
                    FirstRow = dimRecord.Range.FromRow;
                    LastRow = dimRecord.Range.ToRow + 1;
                    FirstColumn = dimRecord.Range.FromColumn;
                    LastColumn = dimRecord.Range.ToColumn + 1;
                    break;
            }
        }

        ColumnWidths = columnWidths;
        MergeCells = [.. cellRanges];

        if (rowIndexMaximum != int.MinValue && columnIndexMaximum != int.MinValue)
        {
            FieldCount = columnIndexMaximum + 1;
            RowCount = rowIndexMaximum + 1;
        }
    }

    public int FieldCount { get; }

    public int RowCount { get; }

        public int FirstRow { get; private set; }

        public int LastRow { get; private set; }

        public int FirstColumn { get; private set; }

        public int LastColumn { get; private set; }

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
                var format = Workbook.GetNumberFormatString(numberFormatIndex);
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
                NumberFormatString numberFormat = Workbook.GetNumberFormatString(numberFormatIndex);
                if (numberFormat.IsTimeSpanFormat && TryParseToTimeSpan(s, out var timeSpan))
                {
                    return timeSpan;
                }

                if (numberFormat.IsDateTimeFormat && DateTimeOffset.TryParse(s, out DateTimeOffset dateTimeOffset))
                {
                    return dateTimeOffset;
                }

                return s;

            default:
                return value;
        }
    }
}
