using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsByteString : IXlsString
    {
        private readonly byte[] m_bytes;
        private readonly uint m_offset;
        private readonly Encoding m_encoding;

        public XlsByteString(byte[] bytes, uint offset, Encoding encoding)
        {
            m_bytes = bytes;
            m_offset = offset;
            m_encoding = encoding;
        }
        
        /// <summary>
        /// Count of characters in string
        /// </summary>
        public ushort CharacterCount => BitConverter.ToUInt16(m_bytes, (int)m_offset);

        public uint HeadSize => 0;

        public uint TailSize => 0;

        public bool IsMultiByte => !Helpers.IsSingleByteEncoding(m_encoding);

        /// <summary>
        /// Returns string represented by this instance
        /// </summary>
        public string Value
        {
            get
            {
                var bytes = ReadArray(0x2, CharacterCount * (Helpers.IsSingleByteEncoding(m_encoding) ? 1 : 2));
                return m_encoding.GetString(bytes, 0, bytes.Length);
            }
        }

        public byte[] ReadArray(int offset, int size)
        {
            byte[] tmp = new byte[size];
            Buffer.BlockCopy(m_bytes, offset, tmp, 0, size);
            return tmp;
        }
    }
}