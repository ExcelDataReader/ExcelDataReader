using System;
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
        
        protected XlsBiffRecord(byte[] bytes, uint offset)
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
        
        /// <summary>
        /// Returns record at specified offset
        /// </summary>
        /// <param name="bytes">byte array</param>
        /// <param name="offset">position in array</param>
        /// <returns>The record -or- null.</returns>
        public static XlsBiffRecord GetRecord(byte[] bytes, uint offset, int biffVersion)
        {
            if (offset >= bytes.Length)
                return null;

            uint id = BitConverter.ToUInt16(bytes, (int)offset);
            
            // Console.WriteLine("GetRecord {0}", (BIFFRECORDTYPE)Id);
            switch ((BIFFRECORDTYPE)id)
            {
                case BIFFRECORDTYPE.BOF_V2:
                case BIFFRECORDTYPE.BOF_V3:
                case BIFFRECORDTYPE.BOF_V4:
                case BIFFRECORDTYPE.BOF:
                    return new XlsBiffBOF(bytes, offset);
                case BIFFRECORDTYPE.EOF:
                    return new XlsBiffEof(bytes, offset);
                case BIFFRECORDTYPE.INTERFACEHDR:
                    return new XlsBiffInterfaceHdr(bytes, offset);

                case BIFFRECORDTYPE.SST:
                    return new XlsBiffSST(bytes, offset, biffVersion);

                case BIFFRECORDTYPE.INDEX:
                    return new XlsBiffIndex(bytes, offset, biffVersion == 8);
                case BIFFRECORDTYPE.ROW:
                    return new XlsBiffRow(bytes, offset);
                case BIFFRECORDTYPE.DBCELL:
                    return new XlsBiffDbCell(bytes, offset);

                case BIFFRECORDTYPE.BOOLERR:
                case BIFFRECORDTYPE.BOOLERR_OLD:
                case BIFFRECORDTYPE.BLANK:
                case BIFFRECORDTYPE.BLANK_OLD:
                    return new XlsBiffBlankCell(bytes, offset);
                case BIFFRECORDTYPE.MULBLANK:
                    return new XlsBiffMulBlankCell(bytes, offset);
                case BIFFRECORDTYPE.LABEL_OLD:
                case BIFFRECORDTYPE.LABEL:
                case BIFFRECORDTYPE.RSTRING:
                    return new XlsBiffLabelCell(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.LABELSST:
                    return new XlsBiffLabelSSTCell(bytes, offset);
                case BIFFRECORDTYPE.INTEGER:
                case BIFFRECORDTYPE.INTEGER_OLD:
                    return new XlsBiffIntegerCell(bytes, offset);
                case BIFFRECORDTYPE.NUMBER:
                case BIFFRECORDTYPE.NUMBER_OLD:
                    return new XlsBiffNumberCell(bytes, offset);
                case BIFFRECORDTYPE.RK:
                    return new XlsBiffRKCell(bytes, offset);
                case BIFFRECORDTYPE.MULRK:
                    return new XlsBiffMulRKCell(bytes, offset);
                case BIFFRECORDTYPE.FORMULA:
                case BIFFRECORDTYPE.FORMULA_V3:
                case BIFFRECORDTYPE.FORMULA_V4:
                    return new XlsBiffFormulaCell(bytes, offset);
                case BIFFRECORDTYPE.FORMAT_V23:
                case BIFFRECORDTYPE.FORMAT:
                    return new XlsBiffFormatString(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.STRING:
                case BIFFRECORDTYPE.STRING_OLD:
                    return new XlsBiffFormulaString(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.CONTINUE:
                    return new XlsBiffContinue(bytes, offset);
                case BIFFRECORDTYPE.DIMENSIONS:
                    return new XlsBiffDimensions(bytes, offset, biffVersion == 8);
                case BIFFRECORDTYPE.BOUNDSHEET:
                    return new XlsBiffBoundSheet(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.WINDOW1:
                    return new XlsBiffWindow1(bytes, offset);
                case BIFFRECORDTYPE.CODEPAGE:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.FNGROUPCOUNT:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.RECORD1904:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.BOOKBOOL:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.BACKUP:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.HIDEOBJ:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.USESELFS:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.UNCALCED:
                    return new XlsBiffUncalced(bytes, offset);
                case BIFFRECORDTYPE.QUICKTIP:
                    return new XlsBiffQuickTip(bytes, offset);
                case BIFFRECORDTYPE.MSODRAWING:
                    return new XlsBiffMSODrawing(bytes, offset);
                case BIFFRECORDTYPE.FILEPASS:
                    return new XlsBiffFilePass(bytes, offset);

                default:
                    return new XlsBiffRecord(bytes, offset);
            }
        }

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
