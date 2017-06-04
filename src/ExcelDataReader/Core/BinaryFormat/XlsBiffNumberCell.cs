namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a floating-point number 
    /// </summary>
    internal class XlsBiffNumberCell : XlsBiffBlankCell
    {
        internal XlsBiffNumberCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader)
        {
        }

        /// <summary>
        /// Gets the value of this cell
        /// </summary>
        public double Value => ReadDouble(0x6);
    }
}