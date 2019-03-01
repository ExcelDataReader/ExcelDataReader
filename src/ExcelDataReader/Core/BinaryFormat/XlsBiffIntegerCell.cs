namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a constant integer number in range 0..65535
    /// </summary>
    internal class XlsBiffIntegerCell : XlsBiffBlankCell
    {
        internal XlsBiffIntegerCell(byte[] bytes, uint offset, int biffVersion)
            : base(bytes, offset, biffVersion)
        {
            if (Id == BIFFRECORDTYPE.INTEGER_OLD)
            {
                Value = ReadUInt16(0x7);
            }
            else
            {
                Value = ReadUInt16(0x6);
            }
        }

        /// <summary>
        /// Gets the cell value.
        /// </summary>
        public int Value { get; }
    }
}