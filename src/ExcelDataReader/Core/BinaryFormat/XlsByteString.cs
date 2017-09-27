using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Word-sized string, stored as single bytes with encoding from CodePage record. Used in BIFF2-5 
    /// </summary>
    internal class XlsByteString : IXlsString
    {
        private readonly byte[] _bytes;
        private readonly uint _offset;

        public XlsByteString(byte[] bytes, uint offset)
        {
            _bytes = bytes;
            _offset = offset;
        }
        
        /// <summary>
        /// Gets the number of characters in the string.
        /// </summary>
        public ushort CharacterCount => BitConverter.ToUInt16(_bytes, (int)_offset);

        /// <summary>
        /// Gets the value.
        /// </summary>
        public string GetValue(Encoding encoding)
        {
            var stringBytes = ReadArray(0x2, CharacterCount * (Helpers.IsSingleByteEncoding(encoding) ? 1 : 2));
            return encoding.GetString(stringBytes, 0, stringBytes.Length);
        }

        public byte[] ReadArray(int offset, int size)
        {
            byte[] tmp = new byte[size];
            Buffer.BlockCopy(_bytes, (int)(_offset + offset), tmp, 0, size);
            return tmp;
        }
    }
}