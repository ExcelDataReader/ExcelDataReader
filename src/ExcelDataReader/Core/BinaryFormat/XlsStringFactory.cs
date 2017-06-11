using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsStringFactory
    {
        public static IXlsString CreateXlsString(byte[] bytes, uint offset, bool isV8, Encoding encoding)
        {
            if (isV8)
                return new XlsFormattedUnicodeString(bytes, offset);

            return new XlsByteString(bytes, offset, encoding);
        }
    }
}