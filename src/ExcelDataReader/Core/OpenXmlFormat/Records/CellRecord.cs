namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class CellRecord(int columnIndex, int xfIndex, object value, CellError? error) : Record
{
    public int ColumnIndex { get; } = columnIndex;

    public int XfIndex { get; } = xfIndex;

    public object Value { get; } = value;

    public CellError? Error { get; } = error;
}
