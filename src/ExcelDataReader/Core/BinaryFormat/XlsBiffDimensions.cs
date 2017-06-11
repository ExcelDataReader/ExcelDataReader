namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Dimensions of worksheet
    /// </summary>
    internal class XlsBiffDimensions : XlsBiffRecord
    {
        internal XlsBiffDimensions(byte[] bytes, uint offset, bool isV8)
            : base(bytes, offset)
        {
            IsV8 = isV8;
        }

        /// <summary>
        /// Gets the index of first row.
        /// </summary>
        public uint FirstRow => IsV8 ? ReadUInt32(0x0) : ReadUInt16(0x0);

        /// <summary>
        /// Gets the index of last row + 1.
        /// </summary>
        public uint LastRow => IsV8 ? ReadUInt32(0x4) : ReadUInt16(0x2);

        /// <summary>
        /// Gets the index of first column.
        /// </summary>
        public ushort FirstColumn => IsV8 ? ReadUInt16(0x8) : ReadUInt16(0x4);

        /// <summary>
        /// Gets the index of last column + 1.
        /// </summary>
        public ushort LastColumn => IsV8 ? (ushort)((ReadUInt16(0x9) >> 8) + 1) : ReadUInt16(0x6);

        /// <summary>
        /// Gets a value indicating whether BIFF8 addressing is used or not.
        /// </summary>
        private bool IsV8 { get; }
    }
}