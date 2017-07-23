using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsBiffCodeName : XlsBiffRecord
    {
        private readonly IXlsString _xlsString;

        internal XlsBiffCodeName(byte[] bytes, uint offset)
            : base(bytes, offset)
        {
            // BIFF8 only
            _xlsString = new XlsUnicodeString(bytes, offset + 4 + 0);
        }

        public string GetValue(Encoding encoding)
        {
            return _xlsString.GetValue(encoding);
        }
    }
}
