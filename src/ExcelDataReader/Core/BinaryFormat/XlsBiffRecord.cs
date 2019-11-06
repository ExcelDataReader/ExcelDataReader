using System;
using System.IO;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents basic BIFF record
    /// Base class for all BIFF record types
    /// </summary>
    internal class XlsBiffRecord
    {
        protected const int ContentOffset = 4;
        
        public XlsBiffRecord(byte[] bytes)
        {
            if (bytes.Length < 4)
                throw new ArgumentException(Errors.ErrorBiffRecordSize);
            Bytes = bytes;
        }
        
        /// <summary>
        /// Gets the type Id of this entry
        /// </summary>
        public BIFFRECORDTYPE Id => (BIFFRECORDTYPE)BitConverter.ToUInt16(Bytes, 0);

        /// <summary>
        /// Gets the data size of this entry
        /// </summary>
        public ushort RecordSize => BitConverter.ToUInt16(Bytes, 2);

        /// <summary>
        /// Gets the whole size of structure
        /// </summary>
        public int Size => ContentOffset + RecordSize;
        
        internal byte[] Bytes { get; }

        public byte ReadByte(int offset)
        {
            return Buffer.GetByte(Bytes, ContentOffset + offset);
        }

        public ushort ReadUInt16(int offset)
        {
            return BitConverter.ToUInt16(Bytes, ContentOffset + offset);
        }

        public uint ReadUInt32(int offset)
        {
            return BitConverter.ToUInt32(Bytes, ContentOffset + offset);
        }

        public ulong ReadUInt64(int offset)
        {
            return BitConverter.ToUInt64(Bytes, ContentOffset + offset);
        }

        public short ReadInt16(int offset)
        {
            return BitConverter.ToInt16(Bytes, ContentOffset + offset);
        }

        public int ReadInt32(int offset)
        {
            return BitConverter.ToInt32(Bytes, ContentOffset + offset);
        }

        public long ReadInt64(int offset)
        {
            return BitConverter.ToInt64(Bytes, ContentOffset + offset);
        }

        public byte[] ReadArray(int offset, int size)
        {
            byte[] tmp = new byte[size];
            Buffer.BlockCopy(Bytes, ContentOffset + offset, tmp, 0, size);
            return tmp;
        }

        public float ReadFloat(int offset)
        {
            return BitConverter.ToSingle(Bytes, ContentOffset + offset);
        }

        public double ReadDouble(int offset)
        {
            return BitConverter.ToDouble(Bytes, ContentOffset + offset);
        }
    }
}
