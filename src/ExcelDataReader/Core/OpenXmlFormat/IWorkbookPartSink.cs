using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat;

internal interface IWorkbookPartSink
{
    void SetDate1904(bool date1904);

    void AddSheet(SheetRecord sheet);

    void SetActiveSheet(int activeSheet);
}
