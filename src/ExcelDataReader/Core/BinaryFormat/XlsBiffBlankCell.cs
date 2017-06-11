namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents blank cell
    /// Base class for all cell types
    /// </summary>
    internal class XlsBiffBlankCell : XlsBiffRecord
    {
        internal XlsBiffBlankCell(byte[] bytes, uint offset)
            : base(bytes, offset)
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
        /// Gets the format used for this cell.
        /// </summary>
        public ushort XFormat => ReadUInt16(0x4);

        /// <inheritdoc />
        public override bool IsCell => true;
    }
}