using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// [MS-XLS] 2.5.240 ShortXLUnicodeString
    /// Byte-sized string, stored as single or multibyte unicode characters.
    /// </summary>
    internal class XlsShortUnicodeString : IXlsString
    {
        private readonly byte[] _bytes;
        private readonly uint _offset;

        public XlsShortUnicodeString(byte[] bytes, uint offset)
        {
            _bytes = bytes;
            _offset = offset;
        }

        public ushort CharacterCount => _bytes[_offset];

        /// <summary>
        /// Gets a value indicating whether the string is a multibyte string or not.
        /// </summary>
        public bool IsMultiByte => (_bytes[_offset + 1] & 0x01) != 0;

        public string GetValue(Encoding encoding)
        {
            if (IsMultiByte)
            {
                return Encoding.Unicode.GetString(_bytes, (int)_offset + 2, CharacterCount * 2);
            }

            byte[] bytes = new byte[CharacterCount * 2];
            for (int i = 0; i < CharacterCount; i++)
            {
                bytes[i * 2] = _bytes[_offset + 2 + i];
            }

            return Encoding.Unicode.GetString(bytes);
        }
    }
}
