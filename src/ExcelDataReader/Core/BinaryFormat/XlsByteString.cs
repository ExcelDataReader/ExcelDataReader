using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsByteString : IXlsString
    {
        private readonly byte[] _bytes;
        private readonly uint _offset;
        private readonly Encoding _encoding;

        public XlsByteString(byte[] bytes, uint offset, Encoding encoding)
        {
            this._bytes = bytes;
            this._offset = offset;
            this._encoding = encoding;
        }
        
        /// <summary>
        /// Gets the number of characters in the string.
        /// </summary>
        public ushort CharacterCount => BitConverter.ToUInt16(_bytes, (int)_offset);

        public uint HeadSize => 0;

        public uint TailSize => 0;

        public bool IsMultiByte => !Helpers.IsSingleByteEncoding(_encoding);

        /// <summary>
        /// Gets the value.
        /// </summary>
        public string Value
        {
            get
            {
                var stringBytes = ReadArray(0x2, CharacterCount * (Helpers.IsSingleByteEncoding(_encoding) ? 1 : 2));
                return _encoding.GetString(stringBytes, 0, stringBytes.Length);
            }
        }

        public byte[] ReadArray(int offset, int size)
        {
            byte[] tmp = new byte[size];
            Buffer.BlockCopy(_bytes, (int)(this._offset + offset), tmp, 0, size);
            return tmp;
        }
    }
}