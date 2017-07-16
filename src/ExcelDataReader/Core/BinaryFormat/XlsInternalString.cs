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

        public ushort CharacterCount => throw new NotImplementedException();

        public uint HeadSize => throw new NotImplementedException();

        public uint TailSize => throw new NotImplementedException();

        public bool IsMultiByte => throw new NotImplementedException();

        public string GetValue(Encoding encoding)
        {
            return stringValue;
        }
    }
}
