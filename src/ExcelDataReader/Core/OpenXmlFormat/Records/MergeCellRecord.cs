namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class MergeCellRecord(CellRange range) : Record
{
    public CellRange Range { get; } = range;
}
