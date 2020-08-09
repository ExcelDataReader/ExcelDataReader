using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a cell containing formula
    /// </summary>
    internal class XlsBiffFormulaCell : XlsBiffBlankCell
    {
        // private FormulaFlags _flags;
        private readonly int _biffVersion;
        private bool _booleanValue;
        private CellError _errorValue;
        private double _xNumValue;
        private FormulaValueType _formulaType;
        private bool _initialized;

        internal XlsBiffFormulaCell(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            _biffVersion = biffVersion;
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
        /// Gets the formula value type.
        /// </summary>
        public FormulaValueType FormulaType
        {
            get
            {
                LazyInit();
                return _formulaType;
            }
        }

        public bool BooleanValue
        {
            get
            {
                LazyInit();
                return _booleanValue;
            }
        }

        public CellError ErrorValue
        {
            get
            {
                LazyInit();
                return _errorValue;
            }
        }

        public double XNumValue
        {
            get
            {
                LazyInit();
                return _xNumValue;
            }
        }

        /*
        public FormulaFlags Flags
        {
            get
            {
                LazyInit();
                return _flags;
            }
        }
        */

        private void LazyInit()
        {
            if (_initialized)
                return;
            _initialized = true;

            if (_biffVersion == 2)
            {
                // _flags = (FormulaFlags)ReadUInt16(0xF);
                _xNumValue = ReadDouble(0x7);
                _formulaType = FormulaValueType.Number;
            }
            else
            {
                // _flags = (FormulaFlags)ReadUInt16(0xE);
                var formulaValueExprO = ReadUInt16(0xC);
                if (formulaValueExprO != 0xFFFF)
                {
                    _formulaType = FormulaValueType.Number;
                    _xNumValue = ReadDouble(0x6);
                }
                else
                {
                    // var formulaLength = ReadByte(0xF);
                    var formulaValueByte1 = ReadByte(0x6);
                    var formulaValueByte3 = ReadByte(0x8);
                    switch (formulaValueByte1)
                    {
                        case 0x00:
                            _formulaType = FormulaValueType.String;
                            break;
                        case 0x01:
                            _formulaType = FormulaValueType.Boolean;
                            _booleanValue = formulaValueByte3 != 0;
                            break;
                        case 0x02:
                            _formulaType = FormulaValueType.Error;
                            _errorValue = (CellError)formulaValueByte3;
                            break;
                        case 0x03:
                            _formulaType = FormulaValueType.EmptyString;
                            break;
                        default:
                            _formulaType = FormulaValueType.Unknown;
                            break;
                    }
                }
            }
        }
    }
}