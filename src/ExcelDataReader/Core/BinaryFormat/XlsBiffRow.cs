namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents row record in table
    /// </summary>
    internal class XlsBiffRow : XlsBiffRecord
    {
        internal XlsBiffRow(byte[] bytes, uint offset)
            : base(bytes, offset)
        {
        }

        /// <summary>
        /// Gets the zero-based index of row described
        /// </summary>
        public ushort RowIndex => ReadUInt16(0x0);

        /// <summary>
        /// Gets the index of first defined column
        /// </summary>
        public ushort FirstDefinedColumn => ReadUInt16(0x2);

        /// <summary>
        /// Gets the index of last defined column
        /// </summary>
        public ushort LastDefinedColumn => ReadUInt16(0x4);

        /// <summary>
        /// Gets the row height
        /// </summary>
        public uint RowHeight => ReadUInt16(0x6);

        /// <summary>
        /// Gets the row flags
        /// </summary>
        public ushort Flags => ReadUInt16(0xC);

        /// <summary>
        /// Gets the default format for this row
        /// </summary>
        public ushort XFormat => ReadUInt16(0xE);
    }
}