using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

internal interface IXlsWorkbookPartSink
{
    Encoding Encoding { get; }

    void AddSheet(XlsSheetInfo sheet);

    void AddFont(XlsFontInfo font);

    void SetCodePage(ushort codePage);

    void SetDate1904(bool date1904);

    void SetActiveSheet(int activeSheet);

    void SetSharedStringTable();
}