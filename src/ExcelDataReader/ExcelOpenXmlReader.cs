using ExcelDataReader.Core.OpenXmlFormat;

namespace ExcelDataReader;

internal sealed class ExcelOpenXmlReader : ExcelDataReader<XlsxWorkbook, XlsxWorksheet>
{
    public ExcelOpenXmlReader(Stream stream, bool singlePassMode = false)
    {
        Document = new(stream);
        Workbook = new XlsxWorkbook(Document);
        Workbook.SinglePassMode = singlePassMode;
        SinglePassMode = singlePassMode;

        // By default, the data reader is positioned on the first result.
        Reset();
    }

    private ZipWorker Document { get; set; }

    public override void Close()
    {
        base.Close();

        Document?.Dispose();
        Workbook = null;
        Document = null;
    }
}
