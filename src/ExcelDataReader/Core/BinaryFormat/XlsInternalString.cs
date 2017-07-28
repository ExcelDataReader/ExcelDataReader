using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Plain string without backing storage. Used internally
    /// </summary>
    internal class XlsInternalString : IXlsString
    {
        private readonly string stringValue;

        public XlsInternalString(string value)
        {
            stringValue = value;
        }

        public string GetValue(Encoding encoding)
        {
            return stringValue;
        }
    }
}
