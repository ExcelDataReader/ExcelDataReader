namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents blank cell
    /// Base class for all cell types
    /// </summary>
    internal class XlsBiffBlankCell : XlsBiffRecord
    {
        internal XlsBiffBlankCell(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the zero-based index of row containing this cell.
        /// </summary>
        public ushort RowIndex => ReadUInt16(0x0);

        /// <summary>
        /// Gets the zero-based index of column containing this cell.
        /// </summary>
        public ushort ColumnIndex => ReadUInt16(0x2);

        /// <summary>
        /// Gets the extended format used for this cell. If BIFF2 and this value is 63, this record was preceded by an IXFE record containing the actual XFormat >= 63.
        /// </summary>
        public ushort XFormat => IsBiff2Cell ? (ushort)(ReadByte(0x4) & 0x3F) : ReadUInt16(0x4);

        /// <summary>
        /// Gets the number format used for this cell. Only used in BIFF2 without XF records. Used by Excel 2.0/2.1 instead of XF/IXFE records.
        /// </summary>
        public ushort Format => IsBiff2Cell ? (ushort)(ReadByte(0x5) & 0x3F) : (ushort)0;

        /// <summary>
        /// Gets a value indicating whether the cell's record identifier is BIFF2-specific. 
        /// The shared binary layout of BIFF2 cells are different from BIFF3+.
        /// </summary>
        public bool IsBiff2Cell
        {
            get
            {
                switch (Id)
                {
                    case BIFFRECORDTYPE.NUMBER_OLD:
                    case BIFFRECORDTYPE.INTEGER_OLD:
                    case BIFFRECORDTYPE.LABEL_OLD:
                    case BIFFRECORDTYPE.BLANK_OLD:
                    case BIFFRECORDTYPE.BOOLERR_OLD:
                        return true;
                    default:
                        return false;
                }
            }
        }
    }
}