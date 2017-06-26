using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a cell containing formula
    /// </summary>
    internal class XlsBiffFormulaCell : XlsBiffNumberCell
    {
        internal XlsBiffFormulaCell(byte[] bytes, uint offset)
            : base(bytes, offset)
        {
        }

        [Flags]
        public enum FormulaFlags : ushort
        {
            AlwaysCalc = 0x0001,
            CalcOnLoad = 0x0002,
            SharedFormulaGroup = 0x0008
        }

        public enum FormulaValueType
        {
            Unknown,

            /// <summary>
            /// Indicates that a string value is stored in a String record that immediately follows this record. See[MS - XLS] 2.5.133 FormulaValue.
            /// </summary>
            String, 

            /// <summary>
            /// Indecates that the formula value is an empty string.
            /// </summary>
            EmptyString,

            /// <summary>
            /// Indicates that the <see cref="BooleanValue"/> property is valid.
            /// </summary>
            Boolean,

            /// <summary>
            /// Indicates that the <see cref="ErrorValue"/> property is valid.
            /// </summary>
            Error,

            /// <summary>
            /// Indicates that the <see cref="XNumValue"/> property is valid.
            /// </summary>
            Number
        }

        /// <summary>
        /// Gets the formula flags
        /// </summary>
        public FormulaFlags Flags => (FormulaFlags)ReadUInt16(0xE);

        /// <summary>
        /// Gets the formula value type.
        /// </summary>
        public FormulaValueType FormulaType
        {
            get
            {
                if (FormulaValueExprO != 0xFFFF)
                    return FormulaValueType.Number;

                switch (FormulaValueByte1)
                {
                    case 0x00: return FormulaValueType.String;
                    case 0x01: return FormulaValueType.Boolean;
                    case 0x02: return FormulaValueType.Error;
                    case 0x03: return FormulaValueType.EmptyString;
                    default: return FormulaValueType.Unknown;
                }
            }
        }

        /// <summary>
        /// Gets the formula string length.
        /// </summary>
        public byte FormulaLength => ReadByte(0xF);

        public byte FormulaValueByte1 => ReadByte(0x6);

        public byte FormulaValueByte2 => ReadByte(0x7);

        public byte FormulaValueByte3 => ReadByte(0x8);

        public byte FormulaValueByte4 => ReadByte(0x9);

        public byte FormulaValueByte5 => ReadByte(0xA);

        public byte FormulaValueByte6 => ReadByte(0xB);

        public ushort FormulaValueExprO => ReadUInt16(0xC);

        public bool BooleanValue => FormulaValueByte3 != 0;

        public FORMULAERROR ErrorValue => (FORMULAERROR)FormulaValueByte3;

        public double XNumValue => ReadDouble(0x6);

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