namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents multiple Blank cell
    /// </summary>
    internal class XlsBiffMulBlankCell : XlsBiffBlankCell
    {
        internal XlsBiffMulBlankCell(byte[] bytes, uint offset, int biffVersion)
            : base(bytes, offset, biffVersion)
        {
        }

        /// <summary>
        /// Gets the zero-based index of last described column
        /// </summary>
        public ushort LastColumnIndex => ReadUInt16(RecordSize - 2);

        /// <summary>
        /// Returns format forspecified column, column must be between ColumnIndex and LastColumnIndex
        /// </summary>
        /// <param name="columnIdx">Index of column</param>
        /// <returns>Format</returns>
        public ushort GetXF(ushort columnIdx)
        {
            int ofs = 4 + 6 * (columnIdx - ColumnIndex);
            if (ofs > RecordSize - 2)
                return 0;
            return ReadUInt16(ofs);
        }
    }
}