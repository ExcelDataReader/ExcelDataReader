namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class SheetPrRecord(string codeName) : Record
{
    public string CodeName { get; } = codeName;
}
