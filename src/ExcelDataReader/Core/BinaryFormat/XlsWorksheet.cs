using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents Worksheet section in workbook.
/// </summary>
internal sealed class XlsWorksheet : IWorksheet
{
    private readonly XlsWorksheetMetadata _metadata;
    private Dictionary<long, CellCommentThread>? _comments;

    public XlsWorksheet(XlsWorkbook workbook, XlsSheetInfo refSheet, Stream stream, bool singlePassMode = false)
    {
        Workbook = workbook;
        Stream = stream;
        Name = refSheet.GetSheetName(workbook.Encoding);
        DataOffset = refSheet.StartOffset;
        VisibleState = refSheet.VisibleState switch
        {
            XlsSheetInfo.SheetVisibility.Hidden => "hidden",
            XlsSheetInfo.SheetVisibility.VeryHidden => "veryhidden",
            _ => "visible",
        };
        SinglePassMode = singlePassMode;
        _metadata = new XlsWorksheetPartReader(workbook, refSheet, stream, singlePassMode).Load();
    }

    public string Name { get; }

    public string? CodeName => _metadata.CodeName;

    public string VisibleState { get; }

    public HeaderFooter? HeaderFooter => _metadata.HeaderFooter;

    public CellRange[] MergeCells => _metadata.MergeCells;

    public List<Column> ColumnWidths => _metadata.ColumnWidths;

    public uint DataOffset { get; }

    public Stream Stream { get; }

    public Encoding Encoding => _metadata.Encoding;

    // BIFF stores heights in twips; the common worksheet boundary uses points.
    public double DefaultRowHeight => _metadata.DefaultRowHeight;

    double IWorksheet.DefaultRowHeight => DefaultRowHeight / 20.0;

    public Dictionary<int, XlsRowOffset> RowOffsetMap => _metadata.RowOffsetMap;

    public int FieldCount => _metadata.FieldCount;

    public int RowCount => _metadata.RowCount;

    public CellRange? Dimension => _metadata.Dimension;

    public bool IsDate1904 => _metadata.IsDate1904;

    public bool SinglePassMode { get; }

    public XlsWorkbook Workbook { get; }

    public IWorksheetReader OpenReader() => new XlsWorksheetReader(this);

    public Dictionary<long, CellCommentThread>? LoadComments()
    {
        if (_comments != null)
            return _comments;
        long position = Stream.Position;
        Dictionary<long, CellCommentThread> comments;
        try
        {
            comments = XlsCommentsReader.Load(this);
        }
        finally
        {
            Stream.Position = position;
        }

        return _comments = comments;
    }
}