using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat;

internal sealed class XlsxWorksheet : IWorksheet
{
    private readonly OpenXmlRowOrder.Region[] _orderRegions = [];

    public XlsxWorksheet(ZipWorker document, XlsxWorkbook workbook, SheetRecord refSheet, bool singlePassMode = false)
    {
        Document = document;
        Workbook = workbook;

        Name = refSheet.Name;
        VisibleState = refSheet.VisibleState;
        Path = refSheet.Path;
        DefaultRowHeight = 15;

        if (Path is not { Length: > 0 } worksheetPath)
        {
            ColumnWidths = [];
            return;
        }

        using var sheetStream = Document.GetWorksheetReader(worksheetPath, singlePassMode ? WorksheetReadMode.Materialize : WorksheetReadMode.Scan, workbook);

        if (sheetStream == null)
        {
            ColumnWidths = [];
            return;
        }

        var metadata = new MetadataConsumer(singlePassMode);
        if (singlePassMode)
        {
            sheetStream.Consume(ref metadata);
        }
        else
        {
            while (sheetStream.Consume(ref metadata) == WorksheetReadStatus.Paused)
            {
            }
        }

        ColumnWidths = metadata.ColumnWidths;
        MergeCells = [.. metadata.MergeCells];
        Dimension = metadata.Dimension;
        CodeName = metadata.CodeName;
        HeaderFooter = metadata.HeaderFooter;
        DefaultRowHeight = metadata.DefaultRowHeight;
        _orderRegions = metadata.RowOrder?.Build() ?? [];
        if (singlePassMode)
        {
            // In single-pass mode FieldCount comes from the dimension hint only (may be 0 if absent)
            FieldCount = metadata.MaxColumnIndex != int.MinValue ? metadata.MaxColumnIndex + 1 : 0;
        }
        else if (metadata.MaxRowIndex != int.MinValue && metadata.MaxColumnIndex != int.MinValue)
        {
            FieldCount = metadata.MaxColumnIndex + 1;
            RowCount = metadata.MaxRowIndex + 1;
        }
    }

    public int FieldCount { get; }

    public int RowCount { get; }

    public CellRange? Dimension { get; }

    public string Name { get; }

    public string? CodeName { get; }

    public string VisibleState { get; }

    public HeaderFooter? HeaderFooter { get; }

    public CellRange[] MergeCells { get; } = [];

    public List<Column> ColumnWidths { get; }

    public double DefaultRowHeight { get; }

    private string? Path { get; }

    private ZipWorker Document { get; }

    private XlsxWorkbook Workbook { get; }

    public IWorksheetReader OpenReader() =>
        Path is { Length: > 0 } worksheetPath
            ? Document.GetWorksheetReader(worksheetPath, WorksheetReadMode.Materialize, Workbook) is { } reader
                ? new OpenXmlRowReader(reader, DefaultRowHeight, _orderRegions)
                : EmptyReader.Instance
            : EmptyReader.Instance;

    private struct MetadataConsumer(bool pauseAtSheetData) : IWorksheetHandler
    {
        public OpenXmlRowOrder? RowOrder { get; } = pauseAtSheetData ? null : new();

        public List<Column> ColumnWidths { get; } = [];

        public List<CellRange> MergeCells { get; } = [];

        public CellRange? Dimension { get; private set; }

        public string? CodeName { get; private set; }

        public HeaderFooter? HeaderFooter { get; private set; }

        public double DefaultRowHeight { get; private set; } = 15;

        public int MaxRowIndex { get; private set; } = int.MinValue;

        public int MaxColumnIndex { get; private set; } = int.MinValue;

        public WorksheetReadControl OnDimension(CellRange range)
        {
            Dimension = range;
            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnCodeName(string? codeName)
        {
            CodeName = codeName;
            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnDefaultRowHeight(double? height)
        {
            if (height != null)
                DefaultRowHeight = height.Value;

            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnColumn(Column column)
        {
            ColumnWidths.Add(column);
            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnMergeCell(CellRange range)
        {
            MergeCells.Add(range);
            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnHeaderFooter(HeaderFooter headerFooter)
        {
            HeaderFooter = headerFooter;
            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnSheetDataBegin() =>
            pauseAtSheetData ? WorksheetReadControl.Pause : WorksheetReadControl.Continue;

        public WorksheetReadControl OnScanResult(int maxRowIndex, int maxColumnIndex)
        {
            MaxRowIndex = Math.Max(MaxRowIndex, maxRowIndex);
            MaxColumnIndex = Math.Max(MaxColumnIndex, maxColumnIndex);
            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnRow(int index, bool hidden, double? height)
        {
            RowOrder?.Add(index);
            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnCell(in Cell cell) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnRowEnd() => WorksheetReadControl.Continue;
    }

    private sealed class EmptyReader : IWorksheetReader
    {
        internal static readonly EmptyReader Instance = new();

        public WorksheetReadStatus Consume<THandler>(ref THandler handler)
            where THandler : struct, IWorksheetHandler => WorksheetReadStatus.EndOfInput;

        public void Dispose()
        {
        }
    }
}