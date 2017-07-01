namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Dimensions of worksheet
    /// </summary>
    internal class XlsBiffDimensions : XlsBiffRecord
    {
        internal XlsBiffDimensions(byte[] bytes, uint offset, int biffVersion)
            : base(bytes, offset)
        {
            if (biffVersion < 8)
            {
                FirstRow = ReadUInt16(0x0);
                LastRow = ReadUInt16(0x2);
                FirstColumn = ReadUInt16(0x4);
                LastColumn = ReadUInt16(0x6);
            }
            else
            {
                FirstRow = ReadUInt32(0x0); // TODO: [MS-XLS] RwLongU
                LastRow = ReadUInt32(0x4);
                FirstColumn = ReadUInt16(0x8); // TODO: [MS-XLS] ColU
                LastColumn = ReadUInt16(0xA);
            }
        }

        /// <summary>
        /// Gets the index of first row.
        /// </summary>
        public uint FirstRow { get; }

        /// <summary>
        /// Gets the index of last row + 1.
        /// </summary>
        public uint LastRow { get; }

        /// <summary>
        /// Gets the index of first column.
        /// </summary>
        public ushort FirstColumn { get; }

        /// <summary>
        /// Gets the index of last column + 1.
        /// </summary>
        public ushort LastColumn { get; }
    }
}