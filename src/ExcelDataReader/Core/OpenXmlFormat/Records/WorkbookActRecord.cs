namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class WorkbookActRecord : Record
{
    public WorkbookActRecord(int activeSheet)
    {
        this.ActiveSheet = activeSheet;
    }

    public int ActiveSheet { get; }
}
