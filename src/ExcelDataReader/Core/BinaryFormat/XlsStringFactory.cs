using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsStringFactory
    {
        public static IXlsString CreateXlsString(byte[] bytes, uint offset, ExcelBinaryReader reader)
        {
            if (reader.IsV8())
                return new XlsFormattedUnicodeString(bytes, offset);

            return new XlsByteString(bytes, offset, reader.Encoding);
        }
    }
}   