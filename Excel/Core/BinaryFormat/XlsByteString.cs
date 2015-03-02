using System;
using System.Text;

namespace ExcelDataReader.Portable.Core.BinaryFormat
{
    internal class XlsByteString : IXlsString
    {
        protected byte[] m_bytes;
        protected uint m_offset;
        private readonly Encoding encoding;

        public XlsByteString(byte[] bytes, uint offset, Encoding encoding)
        {
            m_bytes = bytes;
            m_offset = offset;
            this.encoding = encoding;
        }


        /// <summary>
        /// Count of characters in string
        /// </summary>
        public ushort CharacterCount
        {
            get { return BitConverter.ToUInt16(m_bytes, (int)m_offset); }
        }

        public uint HeadSize
        {
            get { return 0; }
        }

        public uint TailSize
        {
            get { return 0; }
        }

        public bool IsMultiByte
        {
            get { return !Helpers.IsSingleByteEncoding(encoding); }
        }

        public Encoding UseEncoding
        {
            get
            {
                return IsMultiByte ? Encoding.Unicode : Encoding.GetEncoding("windows-1250");
            }
        }

        /// <summary>
        /// Returns string represented by this instance
        /// </summary>
        public string Value
        {
            get
            {
                var bytes = 
                    ReadArray(0x2, CharacterCount * (Helpers.IsSingleByteEncoding(encoding) ? 1 : 2));
                return encoding.GetString(bytes, 0, bytes.Length);
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