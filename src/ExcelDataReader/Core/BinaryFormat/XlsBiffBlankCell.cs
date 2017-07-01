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

            if (biffVersion == 2)
            {
                var cellAttribute1 = ReadByte(0x4);
                XFormat = (ushort)(cellAttribute1 & 0x3F);
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
        /// Gets the format used for this cell. If BIFF2 and this value is 63, this record was preceded by an IXFE record containing the actual XFormat >= 63.
        /// </summary>
        public ushort XFormat { get; }

        /// <inheritdoc />
        public override bool IsCell => true;
    }
}