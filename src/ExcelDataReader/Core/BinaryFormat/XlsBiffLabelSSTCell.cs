namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a string stored in SST.
    /// </summary>
    internal sealed class XlsBiffLabelSSTCell : XlsBiffBlankCell
    {
        internal XlsBiffLabelSSTCell(byte[] bytes)
            : base(bytes)
        {
        }

        public override bool IsEmpty => false;

        /// <summary>
        /// Gets the index of string in Shared String Table.
        /// </summary>
        public uint SSTIndex => ReadUInt32(0x6);
    }
}