using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a cell containing formula
    /// </summary>
    internal class XlsBiffFormulaCell : XlsBiffNumberCell
    {
        internal XlsBiffFormulaCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader)
        {
        }

        [Flags]
        public enum FormulaFlags : ushort
        {
            AlwaysCalc = 0x0001,
            CalcOnLoad = 0x0002,
            SharedFormulaGroup = 0x0008
        }

        /// <summary>
        /// Gets the formula flags
        /// </summary>
        public FormulaFlags Flags => (FormulaFlags)ReadUInt16(0xE);

        /// <summary>
        /// Gets the formula string length.
        /// </summary>
        public byte FormulaLength => ReadByte(0xF);

        /// <summary>
        /// Gets the type-dependent value of formula
        /// </summary>
        public new object Value
        {
            get
            {
                long val = ReadInt64(0x6);
                if (((ulong)val & 0xFFFF000000000000) == 0xFFFF000000000000)
                {
                    byte type = (byte)(val & 0xFF);
                    byte code = (byte)((val >> 16) & 0xFF);
                    switch (type)
                    {
                        case 0: // String

                            //////////////fix
                            XlsBiffRecord rec = GetRecord(Bytes, (uint)(Offset + Size), Reader);
                            XlsBiffFormulaString str;
                            if (rec.Id == BIFFRECORDTYPE.SHAREDFMLA)
                                str = GetRecord(Bytes, (uint)(Offset + Size + rec.Size), Reader) as XlsBiffFormulaString;
                            else
                                str = rec as XlsBiffFormulaString;
                            //////////////fix

                            if (str == null)
                                return string.Empty;
                            else
                                return str.Value;

                        case 1: // Boolean

                            return code != 0;
                        case 2: // Error

                            return (FORMULAERROR)code;
                        default:
                            return null;
                    }
                }

                return Helpers.Int64BitsToDouble(val);
            }
        }

        public string Formula
        {
            get
            {
                byte[] bts = ReadArray(0x10, FormulaLength);
                return Encoding.Unicode.GetString(bts, 0, bts.Length);
            }
        }
    }
}