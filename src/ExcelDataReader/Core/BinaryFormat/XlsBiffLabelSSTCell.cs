using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a string stored in SST
    /// </summary>
    internal class XlsBiffLabelSSTCell : XlsBiffBlankCell
    {
        internal XlsBiffLabelSSTCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader)
        {
        }

        /// <summary>
        /// Gets the index of string in Shared String Table
        /// </summary>
        public uint SSTIndex => ReadUInt32(0x6);
    }
}