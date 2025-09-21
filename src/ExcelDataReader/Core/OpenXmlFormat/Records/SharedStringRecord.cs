namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class SharedStringRecord(string value) : Record
{
    public string Value { get; } = value;
}
