using ExcelDataReader.Core;
using ExcelDataReader.Core.BinaryFormat;

namespace ExcelDataReader;

internal sealed class ExcelBinaryReader : ExcelDataReader<XlsWorkbook, XlsWorksheet>
{
    public ExcelBinaryReader(Stream stream, ExcelReaderConfiguration configuration)
    {
        try
        {
            Workbook = new XlsWorkbook(stream, configuration);
            Workbook.SinglePassMode = configuration.SinglePassMode;
            SinglePassMode = configuration.SinglePassMode;

            // By default, the data reader is positioned on the first result.
            Reset();
        }
        catch (Exception exception)
        {
            ResourceCleanup.DisposeAll(exception, this, stream);
            throw;
        }
    }
}
