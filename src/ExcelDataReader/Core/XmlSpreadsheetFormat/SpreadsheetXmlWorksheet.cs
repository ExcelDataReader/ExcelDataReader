#nullable enable

using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core.XmlSpreadsheetFormat;

internal sealed class SpreadsheetXmlWorksheet : IWorksheet
{
    private readonly Stream _stream;
    private readonly IReadOnlyDictionary<string, ExtendedFormat> _stylesById;
    private readonly IReadOnlyDictionary<int, NumberFormatString>? _formats;
    private readonly int _worksheetIndex;
    private Dictionary<long, CellCommentThread>? _comments;
    private HyperlinkStore? _hyperlinks;

    private SpreadsheetXmlWorksheet(
        string name,
        string visibleState,
        string? codeName,
        HeaderFooter? headerFooter,
        List<Column> columnWidths,
        CellRange[] mergeCells,
        int fieldCount,
        int rowCount,
        CellRange? dimension,
        double defaultRowHeight,
        Stream stream,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
        IReadOnlyDictionary<int, NumberFormatString>? formats,
        int worksheetIndex)
    {
        Name = name;
        CodeName = codeName;
        VisibleState = visibleState;
        HeaderFooter = headerFooter;
        ColumnWidths = columnWidths;
        MergeCells = mergeCells;
        FieldCount = fieldCount;
        RowCount = rowCount;
        Dimension = dimension;
        DefaultRowHeight = defaultRowHeight;
        _stream = stream;
        _stylesById = stylesById;
        _formats = formats;
        _worksheetIndex = worksheetIndex;
    }

    public string Name { get; }

    public string? CodeName { get; }

    public string VisibleState { get; }

    public HeaderFooter? HeaderFooter { get; }

    public int FieldCount { get; }

    public int RowCount { get; }

    public CellRange? Dimension { get; }

    public CellRange[] MergeCells { get; }

    public List<Column> ColumnWidths { get; }

    public double DefaultRowHeight { get; }

    public static SpreadsheetXmlWorksheet Create(
        Stream stream,
        int worksheetIndex,
        string name,
        string visibleState,
        string? codeName,
        HeaderFooter? headerFooter,
        int expandedColumnCount,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
        IReadOnlyDictionary<int, NumberFormatString>? formats,
        bool singlePassMode)
    {
        if (singlePassMode)
        {
            return new SpreadsheetXmlWorksheet(
                name,
                visibleState,
                codeName,
                headerFooter,
                [],
                [],
                0,
                0,
                null,
                15D,
                stream,
                stylesById,
                formats,
                worksheetIndex);
        }

        var metadata = new MetadataSink(name, visibleState, codeName, headerFooter);
        using (var reader = new SpreadsheetXmlWorksheetReader(
            stream, worksheetIndex, stylesById, formats, WorksheetReadMode.Scan))
        {
            while (reader.Consume(ref metadata) == WorksheetReadStatus.Paused)
            {
            }
        }

        int fieldCount = metadata.MaxColumnIndex + 1;
        int rowCount = metadata.MaxRowIndex + 1;
        var dimension = fieldCount > 0 && rowCount > 0
            ? new CellRange(0, 0, fieldCount - 1, rowCount - 1)
            : null;

        return new SpreadsheetXmlWorksheet(
            metadata.Name,
            metadata.VisibleState,
            metadata.CodeName,
            metadata.HeaderFooter,
            metadata.ColumnWidths,
            [.. metadata.MergeCells],
            fieldCount,
            rowCount,
            dimension,
            metadata.DefaultRowHeight,
            stream,
            stylesById,
            formats,
            worksheetIndex);
    }

    public IWorksheetReader OpenReader()
        => new SpreadsheetXmlWorksheetReader(_stream, _worksheetIndex, _stylesById, _formats, WorksheetReadMode.Materialize);

    public Dictionary<long, CellCommentThread>? LoadComments()
    {
        if (_comments != null)
            return _comments;
        long position = _stream.Position;
        Dictionary<long, CellCommentThread> comments;
        Exception? failure = null;
        try
        {
            _stream.Position = 0;
            comments = SpreadsheetXmlCommentsReader.Load(_stream, _worksheetIndex);
        }
        catch (Exception error)
        {
            failure = error;
            throw;
        }
        finally
        {
            ResourceCleanup.RestorePosition(_stream, position, failure);
        }

        return _comments = comments;
    }

    public HyperlinkStore? LoadHyperlinks()
    {
        if (_hyperlinks != null)
            return _hyperlinks;
        long position = _stream.Position;
        HyperlinkStore hyperlinks;
        Exception? failure = null;
        try
        {
            _stream.Position = 0;
            hyperlinks = SpreadsheetXmlHyperlinksReader.Load(_stream, _worksheetIndex);
        }
        catch (Exception error)
        {
            failure = error;
            throw;
        }
        finally
        {
            ResourceCleanup.RestorePosition(_stream, position, failure);
        }

        return _hyperlinks = hyperlinks;
    }

    private struct MetadataSink(
        string name,
        string visibleState,
        string? codeName,
        HeaderFooter? headerFooter) : IWorksheetHandler
    {
        public string Name { get; } = name;

        public string VisibleState { get; private set; } = visibleState;

        public string? CodeName { get; private set; } = codeName;

        public HeaderFooter? HeaderFooter { get; private set; } = headerFooter;

        public List<Column> ColumnWidths { get; } = [];

        public List<CellRange> MergeCells { get; } = [];

        public int MaxRowIndex { get; private set; } = -1;

        public int MaxColumnIndex { get; private set; } = -1;

        public double DefaultRowHeight { get; private set; } = 15D;

        public readonly WorksheetReadControl OnDimension(CellRange range) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnCodeName(string? codeName)
        {
            CodeName = codeName;
            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnDefaultRowHeight(double? height)
        {
            if (height.HasValue)
                DefaultRowHeight = height.Value;
            return WorksheetReadControl.Continue;
        }

        public readonly WorksheetReadControl OnColumn(Column column)
        {
            ColumnWidths.Add(column);
            return WorksheetReadControl.Continue;
        }

        public readonly WorksheetReadControl OnMergeCell(CellRange range)
        {
            MergeCells.Add(range);
            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnHeaderFooter(HeaderFooter headerFooter)
        {
            HeaderFooter = headerFooter;
            return WorksheetReadControl.Continue;
        }

        public readonly WorksheetReadControl OnSheetDataBegin() => WorksheetReadControl.Continue;

        public WorksheetReadControl OnScanResult(int maxRowIndex, int maxColumnIndex)
        {
            MaxRowIndex = Math.Max(MaxRowIndex, maxRowIndex);
            MaxColumnIndex = Math.Max(MaxColumnIndex, maxColumnIndex);
            return WorksheetReadControl.Continue;
        }

        public readonly WorksheetReadControl OnRow(int index, bool hidden, double? height) => WorksheetReadControl.Continue;

        public readonly WorksheetReadControl OnCell(in Cell cell) => WorksheetReadControl.Continue;

        public readonly WorksheetReadControl OnRowEnd() => WorksheetReadControl.Continue;
    }
}