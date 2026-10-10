using ExcelDataReader.Core;
using ExcelDataReader.Core.OpenXmlFormat;

namespace ExcelDataReader;

internal sealed class ExcelOpenXmlReader : ExcelDataReader<XlsxWorkbook, XlsxWorksheet>
{
    public ExcelOpenXmlReader(Stream stream, ExcelReaderConfiguration configuration)
    {
        Document = new(stream);
        try
        {
            Workbook = new XlsxWorkbook(Document, configuration);
            Workbook.SinglePassMode = configuration.SinglePassMode;
            SinglePassMode = configuration.SinglePassMode;

            // By default, the data reader is positioned on the first result.
            Reset();
        }
        catch (Exception exception)
        {
            ResourceCleanup.DisposeAll(exception, this, Document);
            throw;
        }
    }

    private ZipWorker? Document { get; set; }
}
