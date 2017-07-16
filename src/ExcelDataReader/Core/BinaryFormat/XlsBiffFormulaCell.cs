using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a cell containing formula
    /// </summary>
    internal class XlsBiffFormulaCell : XlsBiffBlankCell
    {
        internal XlsBiffFormulaCell(byte[] bytes, uint offset, int biffVersion)
            : base(bytes, offset, biffVersion)
        {
            if (biffVersion == 2)
            {
                Flags = (FormulaFlags)ReadUInt16(0xF);
                XNumValue = ReadDouble(0x7);
                FormulaType = FormulaValueType.Number;
            }
            else
            {
                Flags = (FormulaFlags)ReadUInt16(0xE);

                var formulaValueExprO = ReadUInt16(0xC);
                if (formulaValueExprO != 0xFFFF)
                {
                    FormulaType = FormulaValueType.Number;
                    XNumValue = ReadDouble(0x6);
                }
                else
                {
                    var formulaValueByte1 = ReadByte(0x6);
                    var formulaValueByte3 = ReadByte(0x8);
                    var formulaLength = ReadByte(0xF);
                    switch (formulaValueByte1)
                    {
                        case 0x00:
                            FormulaType = FormulaValueType.String;
                            break;
                        case 0x01:
                            FormulaType = FormulaValueType.Boolean;
                            BooleanValue = formulaValueByte3 != 0;
                            break;
                        case 0x02:
                            FormulaType = FormulaValueType.Error;
                            ErrorValue = (FORMULAERROR)formulaValueByte3;
                            break;
                        case 0x03:
                            FormulaType = FormulaValueType.EmptyString;
                            break;
                        default:
                            FormulaType = FormulaValueType.Unknown;
                            break;
                    }
                }
            }
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
        public FormulaFlags Flags { get; }

        /// <summary>
        /// Gets the formula value type.
        /// </summary>
        public FormulaValueType FormulaType { get; }

        public bool BooleanValue { get; }

        public FORMULAERROR ErrorValue { get; }

        public double XNumValue { get; }
    }
}