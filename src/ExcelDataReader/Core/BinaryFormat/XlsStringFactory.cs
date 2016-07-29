using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsStringFactory
    {
        public static IXlsString CreateXlsString(byte[] bytes, uint offset, ExcelBinaryReader reader)
        {
            if (reader.isV8())
                //return new XlsFormattedUnicodeString(bytes, offset, reader.Encoding);
                return new XlsFormattedUnicodeString(bytes, offset);
            else
                return new XlsByteString(bytes, offset, reader.Encoding);
        }
    }
}   