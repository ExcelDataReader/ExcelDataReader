namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class ColumnRecord(Column column) : Record
{
    public Column Column { get; } = column;
}
