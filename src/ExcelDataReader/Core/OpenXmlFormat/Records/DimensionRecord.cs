namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class DimensionRecord(int lastColumn) : Record
{
    public int LastColumn { get; } = lastColumn;
}
