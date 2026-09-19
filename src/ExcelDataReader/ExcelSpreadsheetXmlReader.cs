using ExcelDataReader.Core.XmlSpreadsheetFormat;

namespace ExcelDataReader;

internal sealed class ExcelSpreadsheetXmlReader : ExcelDataReader<SpreadsheetXmlWorkbook, SpreadsheetXmlWorksheet>
{
    private Stream _stream;

    public ExcelSpreadsheetXmlReader(Stream stream, bool singlePassMode = false)
    {
        _stream = stream;
        Workbook = new SpreadsheetXmlWorkbook(stream);
        Workbook.SinglePassMode = singlePassMode;
        SinglePassMode = singlePassMode;

        // By default, the data reader is positioned on the first result.
        Reset();
    }

    public override void Close()
    {
        base.Close();
        _stream?.Dispose();
        _stream = null;
        Workbook = null;
    }
}
