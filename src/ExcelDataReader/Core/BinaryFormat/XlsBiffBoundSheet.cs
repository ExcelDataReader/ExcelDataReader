using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Sheet record in Workbook Globals
    /// </summary>
    internal class XlsBiffBoundSheet : XlsBiffRecord
    {
        private readonly IXlsString _sheetName;

        internal XlsBiffBoundSheet(byte[] bytes, uint offset, int biffVersion)
            : base(bytes, offset)
        {
            StartOffset = ReadUInt32(0x0);
            Type = (SheetType)ReadByte(0x5);
            VisibleState = (SheetVisibility)ReadByte(0x4);

            if (biffVersion == 8)
            {
                _sheetName = new XlsShortUnicodeString(bytes, offset + 4 + 6);
            }
            else if (biffVersion == 5)
            {
                _sheetName = new XlsShortByteString(bytes, offset + 4 + 6);
            }
            else 
            {
                throw new ArgumentException("Unexpected BIFF version " + biffVersion.ToString(), nameof(biffVersion));
            }
        }

        internal XlsBiffBoundSheet(uint startOffset, SheetType type, SheetVisibility visibleState, string name)
            : base(new byte[32], 0)
        {
            StartOffset = startOffset;
            Type = type;
            VisibleState = visibleState;
            _sheetName = new XlsInternalString(name);
        }

        public enum SheetType : byte
        {
            Worksheet = 0x0,
            MacroSheet = 0x1,
            Chart = 0x2,

            // ReSharper disable once InconsistentNaming
            VBModule = 0x6
        }

        public enum SheetVisibility : byte
        {
            Visible = 0x0,
            Hidden = 0x1,
            VeryHidden = 0x2
        }

        /// <summary>
        /// Gets the worksheet data start offset.
        /// </summary>
        public uint StartOffset { get; }

        /// <summary>
        /// Gets the worksheet type.
        /// </summary>
        public SheetType Type { get; }

        /// <summary>
        /// Gets the visibility of the worksheet.
        /// </summary>
        public SheetVisibility VisibleState { get; }

        /// <summary>
        /// Gets the name of the worksheet.
        /// </summary>
        public string GetSheetName(Encoding encoding)
        {
            return _sheetName.GetValue(encoding);
        }
    }
}