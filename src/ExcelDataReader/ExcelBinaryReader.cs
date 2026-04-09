using System.Globalization;
using System.Text;
using ExcelDataReader.Core.BinaryFormat;

namespace ExcelDataReader;

internal sealed class ExcelBinaryReader : ExcelDataReader<XlsWorkbook, XlsWorksheet>
{
    public ExcelBinaryReader(Stream stream, string password, Encoding fallbackEncoding, CultureInfo culture = null, bool singlePassMode = false)
    {
        Workbook = new XlsWorkbook(stream, password, fallbackEncoding);
        Workbook.Culture = culture;
        Workbook.SinglePassMode = singlePassMode;
        SinglePassMode = singlePassMode;
        
        // By default, the data reader is positioned on the first result.
        Reset();
    }

    public override void Close()
    {
        base.Close();
        Workbook?.Stream?.Dispose();
        Workbook = null;
    }
}
