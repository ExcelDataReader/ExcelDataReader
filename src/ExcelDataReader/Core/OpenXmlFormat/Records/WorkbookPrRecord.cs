namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class WorkbookPrRecord(bool date1904) : Record
{
    public bool Date1904 { get; } = date1904;
}