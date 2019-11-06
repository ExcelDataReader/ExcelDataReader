namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents row record in table
    /// </summary>
    internal class XlsBiffRow : XlsBiffRecord
    {
        internal XlsBiffRow(byte[] bytes)
            : base(bytes)
        {
            if (Id == BIFFRECORDTYPE.ROW_V2)
            {
                RowIndex = ReadUInt16(0x0);
                FirstDefinedColumn = ReadUInt16(0x2);
                LastDefinedColumn = ReadUInt16(0x4);
                var heightBits = ReadUInt16(0x6);

                UseDefaultRowHeight = (heightBits & 0x8000) != 0;
                RowHeight = heightBits & 0x7FFFF;

                UseXFormat = ReadByte(0xA) != 0;
                if (UseXFormat)
                    XFormat = ReadUInt16(0x10);
            }
            else
            {
                RowIndex = ReadUInt16(0x0);
                FirstDefinedColumn = ReadUInt16(0x2);
                LastDefinedColumn = ReadUInt16(0x4);

                var heightBits = ReadUInt16(0x6);
                UseDefaultRowHeight = (heightBits & 0x8000) != 0;
                RowHeight = heightBits & 0x7FFFF;

                var flags = (RowHeightFlags)ReadUInt16(0xC);
                RowHeight = (flags & RowHeightFlags.DyZero) == 0 ? RowHeight : 0;

                UseXFormat = (flags & RowHeightFlags.GhostDirty) != 0;
                XFormat = (ushort)(ReadUInt16(0xE) & 0xFFF);
            }
        }

        internal enum RowHeightFlags : ushort
        {
            OutlineLevelMask = 3,
            Collapsed = 16,
            DyZero = 32,
            Unsynced = 64,
            GhostDirty = 128
        }

        /// <summary>
        /// Gets the zero-based index of row described
        /// </summary>
        public ushort RowIndex { get; }

        /// <summary>
        /// Gets the index of first defined column
        /// </summary>
        public ushort FirstDefinedColumn { get; }

        /// <summary>
        /// Gets the index of last defined column
        /// </summary>
        public ushort LastDefinedColumn { get; }

        /// <summary>
        /// Gets a value indicating whether to use the default row height instead of the RowHeight property
        /// </summary>
        public bool UseDefaultRowHeight { get; }

        /// <summary>
        /// Gets the row height in twips.
        /// </summary>
        public int RowHeight { get; }

        /// <summary>
        /// Gets a value indicating whether the XFormat property is used
        /// </summary>
        public bool UseXFormat { get; }

        /// <summary>
        /// Gets the default format for this row
        /// </summary>
        public ushort XFormat { get; }
    }
}