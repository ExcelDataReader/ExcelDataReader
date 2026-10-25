using ExcelDataReader.Core.XmlSpreadsheetFormat;

namespace ExcelDataReader;

internal sealed class ExcelSpreadsheetXmlReader : ExcelDataReader<SpreadsheetXmlWorkbook, SpreadsheetXmlWorksheet>
{
    public ExcelSpreadsheetXmlReader(Stream stream, bool singlePassMode = false)
    {
        Workbook = new SpreadsheetXmlWorkbook(stream);
        DocumentProperties = Workbook.DocumentProperties;
        Workbook.SinglePassMode = singlePassMode;
        SinglePassMode = singlePassMode;

        // By default, the data reader is positioned on the first result.
        Reset();
    }
}
