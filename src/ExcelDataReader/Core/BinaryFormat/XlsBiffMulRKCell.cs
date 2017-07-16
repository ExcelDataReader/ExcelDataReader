namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents multiple RK number cells
    /// </summary>
    internal class XlsBiffMulRKCell : XlsBiffBlankCell
    {
        internal XlsBiffMulRKCell(byte[] bytes, uint offset, int biffVersion)
            : base(bytes, offset, biffVersion)
        {
        }

        /// <summary>
        /// Gets the zero-based index of last described column
        /// </summary>
        public ushort LastColumnIndex => ReadUInt16(RecordSize - 2);

        /// <summary>
        /// Returns format for specified column
        /// </summary>
        /// <param name="columnIdx">Index of column, must be between ColumnIndex and LastColumnIndex</param>
        /// <returns>The format.</returns>
        public ushort GetXF(ushort columnIdx)
        {
            int ofs = 4 + 6 * (columnIdx - ColumnIndex);
            if (ofs > RecordSize - 2)
                return 0;
            return ReadUInt16(ofs);
        }

        /// <summary>
        /// Gets the value for specified column
        /// </summary>
        /// <param name="columnIdx">Index of column, must be between ColumnIndex and LastColumnIndex</param>
        /// <returns>The value.</returns>
        public double GetValue(ushort columnIdx)
        {
            int ofs = 6 + 6 * (columnIdx - ColumnIndex);
            if (ofs > RecordSize)
                return 0;
            return XlsBiffRKCell.NumFromRK(ReadUInt32(ofs));
        }
    }
}