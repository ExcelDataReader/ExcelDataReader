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

        /// <summary>
        /// Gets the formula flags
        /// </summary>
        public FormulaFlags Flags => (FormulaFlags)ReadUInt16(0xE);

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

        /// <summary>
        /// Gets a value indicating whether a string value is stored in a String record that immediately follows this record. See [MS-XLS] 2.5.133 FormulaValue
        /// </summary>
        public bool IsString => FormulaValueExprO == 0xFFFF && FormulaValueByte1 == 0x00;

        /// <summary>
        /// Gets a value indicating whether the BooleanValue property is valid.
        /// </summary>
        public bool IsBoolean => FormulaValueExprO == 0xFFFF && FormulaValueByte1 == 0x01;

        /// <summary>
        /// Gets a value indicating whether the ErrorValue property is valid.
        /// </summary>
        public bool IsError => FormulaValueExprO == 0xFFFF && FormulaValueByte1 == 0x02;

        /// <summary>
        /// Gets a value indicating whether the XNumValue property is valid.
        /// </summary>
        public bool IsXNum => FormulaValueExprO != 0xFFFF;

        /// <summary>
        /// Gets a value indicating whether the formula value is an empty string.
        /// </summary>
        public bool IsEmptyString => FormulaValueExprO == 0xFFFF && FormulaValueByte1 == 0x03;

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