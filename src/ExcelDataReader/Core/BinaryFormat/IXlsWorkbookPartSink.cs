using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

internal interface IXlsWorkbookPartSink
{
    Encoding Encoding { get; }

    void AddSheet(XlsBiffBoundSheet sheet);

    void AddFont(XlsBiffFont font);

    void AddXf(XlsBiffXF xf);

    void AddNumberFormat(int index, string format);

    void SetCodePage(ushort codePage);

    void SetDate1904(bool date1904);

    void SetActiveSheet(int activeSheet);

    void SetSharedStringTable(XlsBiffSST sharedStrings);
}
