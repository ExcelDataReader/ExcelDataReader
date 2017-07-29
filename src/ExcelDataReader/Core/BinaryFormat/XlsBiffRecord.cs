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
        private const int ContentOffset = 4;
        
        public XlsBiffRecord(byte[] bytes, uint offset)
        {
            if (bytes.Length - offset < 4)
                throw new ArgumentException(Errors.ErrorBiffRecordSize);
            Bytes = bytes;
            RecordContentOffset = (int)(4 + offset);
        }
        
        /// <summary>
        /// Gets the type Id of this entry
        /// </summary>
        public BIFFRECORDTYPE Id => (BIFFRECORDTYPE)BitConverter.ToUInt16(Bytes, RecordContentOffset - ContentOffset);

        /// <summary>
        /// Gets the data size of this entry
        /// </summary>
        public ushort RecordSize => BitConverter.ToUInt16(Bytes, RecordContentOffset - 2);

        /// <summary>
        /// Gets the whole size of structure
        /// </summary>
        public int Size => ContentOffset + RecordSize;

        public virtual bool IsCell => false;

        internal byte[] Bytes { get; }

        internal int RecordContentOffset { get; }

        internal int Offset => RecordContentOffset - ContentOffset;

        public byte ReadByte(int offset)
        {
            return Buffer.GetByte(Bytes, RecordContentOffset + offset);
        }

        public ushort ReadUInt16(int offset)
        {
            return BitConverter.ToUInt16(Bytes, RecordContentOffset + offset);
        }

        public uint ReadUInt32(int offset)
        {
            return BitConverter.ToUInt32(Bytes, RecordContentOffset + offset);
        }

        public ulong ReadUInt64(int offset)
        {
            return BitConverter.ToUInt64(Bytes, RecordContentOffset + offset);
        }

        public short ReadInt16(int offset)
        {
            return BitConverter.ToInt16(Bytes, RecordContentOffset + offset);
        }

        public int ReadInt32(int offset)
        {
            return BitConverter.ToInt32(Bytes, RecordContentOffset + offset);
        }

        public long ReadInt64(int offset)
        {
            return BitConverter.ToInt64(Bytes, RecordContentOffset + offset);
        }

        public byte[] ReadArray(int offset, int size)
        {
            byte[] tmp = new byte[size];
            Buffer.BlockCopy(Bytes, RecordContentOffset + offset, tmp, 0, size);
            return tmp;
        }

        public float ReadFloat(int offset)
        {
            return BitConverter.ToSingle(Bytes, RecordContentOffset + offset);
        }

        public double ReadDouble(int offset)
        {
            return BitConverter.ToDouble(Bytes, RecordContentOffset + offset);
        }
    }
}
