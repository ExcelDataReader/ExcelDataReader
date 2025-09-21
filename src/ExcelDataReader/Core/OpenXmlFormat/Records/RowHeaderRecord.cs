namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class RowHeaderRecord(int rowIndex, bool hidden, double? height) : Record
{
    public int RowIndex { get; } = rowIndex;

    public bool Hidden { get; } = hidden;

    public double? Height { get; } = height;
}
