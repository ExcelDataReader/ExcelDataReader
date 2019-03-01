namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents blank cell
    /// Base class for all cell types
    /// </summary>
    internal class XlsBiffBlankCell : XlsBiffRecord
    {
        internal XlsBiffBlankCell(byte[] bytes, uint offset, int biffVersion)
            : base(bytes, offset)
        {
            RowIndex = ReadUInt16(0x0);
            ColumnIndex = ReadUInt16(0x2);

            if (IsBiff2Cell)
            {
                var cellAttribute1 = ReadByte(0x4);
                var cellAttribute2 = ReadByte(0x5);
                XFormat = (ushort)(cellAttribute1 & 0x3F);
                Format = (byte)(cellAttribute2 & 0x3F);
            }
            else
            {
                XFormat = ReadUInt16(0x4);
            }
        }

        /// <summary>
        /// Gets the zero-based index of row containing this cell.
        /// </summary>
        public ushort RowIndex { get; }

        /// <summary>
        /// Gets the zero-based index of column containing this cell.
        /// </summary>
        public ushort ColumnIndex { get; }

        /// <summary>
        /// Gets the extended format used for this cell. If BIFF2 and this value is 63, this record was preceded by an IXFE record containing the actual XFormat >= 63.
        /// </summary>
        public ushort XFormat { get; }

        /// <summary>
        /// Gets the number format used for this cell. Only used in BIFF2 without XF records. Used by Excel 2.0/2.1 instead of XF/IXFE records.
        /// </summary>
        public ushort Format { get; }

        /// <inheritdoc />
        public override bool IsCell => true;

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