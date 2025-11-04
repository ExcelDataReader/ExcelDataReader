namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class CellRecord(int columnIndex, int xfIndex, string refAttr, object value, CellError? error) : Record
{
    public int ColumnIndex { get; } = columnIndex;

    public int XfIndex { get; } = xfIndex;

    public string RefAttr { get; } = refAttr;

    public object Value { get; } = value;

    public CellError? Error { get; } = error;
}
