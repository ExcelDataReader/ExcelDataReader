using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

internal sealed class XlsWorksheetMetadata
{
    public string? CodeName { get; set; }

    public HeaderFooter? HeaderFooter { get; set; }

    public CellRange[] MergeCells { get; set; } = [];

    public List<Column> ColumnWidths { get; set; } = [];

    public Encoding Encoding { get; set; } = Encoding.Unicode;

    public double DefaultRowHeight { get; set; } = 255;

    public Dictionary<int, XlsRowOffset> RowOffsetMap { get; } = [];

    public int FieldCount { get; set; }

    public int RowCount { get; set; }

    public CellRange? Dimension { get; set; }

    public bool IsDate1904 { get; set; }
}
