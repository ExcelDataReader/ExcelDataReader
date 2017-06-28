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
        private readonly Encoding _encoding;

        public XlsShortByteString(byte[] bytes, uint offset, Encoding encoding)
        {
            _bytes = bytes;
            _offset = offset;
            _encoding = encoding;
        }

        public ushort CharacterCount => _bytes[_offset];

        public bool IsMultiByte => Helpers.IsSingleByteEncoding(_encoding);

        public string Value 
        {
            get
            {
                // Supposedly this is never multibyte, but technically could be
                if (IsMultiByte)
                {
                    return _encoding.GetString(_bytes, (int)_offset + 1, CharacterCount * 2);
                }

                return _encoding.GetString(_bytes, (int)_offset + 1, CharacterCount);
            }
        }

        public uint HeadSize => throw new NotImplementedException();

        public uint TailSize => throw new NotImplementedException();
    }
}
