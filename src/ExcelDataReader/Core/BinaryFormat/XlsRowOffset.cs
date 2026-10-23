namespace ExcelDataReader.Core.BinaryFormat;

internal sealed class XlsRowOffset
{
    public bool UseDefaultRowHeight { get; set; }

    public double? RowHeight { get; set; }

    public int MinCellOffset { get; set; }

    public int MaxCellOffset { get; set; }

    public int MaxOverlapRowIndex { get; set; }
}
