using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Byte sized string, stored as bytes, with encoding from CodePage record. Used in BIFF2-5 .
    /// </summary>
    internal class XlsShortByteString : IXlsString
    {
        private readonly byte[] _bytes;
        private readonly uint _offset;

        public XlsShortByteString(byte[] bytes, uint offset)
        {
            _bytes = bytes;
            _offset = offset;
        }

        public ushort CharacterCount => _bytes[_offset];

        public string GetValue(Encoding encoding)
        {
            // Supposedly this is never multibyte, but technically could be
            if (!Helpers.IsSingleByteEncoding(encoding))
            {
                return encoding.GetString(_bytes, (int)_offset + 1, CharacterCount * 2);
            }

            return encoding.GetString(_bytes, (int)_offset + 1, CharacterCount);
        }
    }
}
