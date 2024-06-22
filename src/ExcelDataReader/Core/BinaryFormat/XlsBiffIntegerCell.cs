namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a constant integer number in range 0..65535.
    /// </summary>
    internal sealed class XlsBiffIntegerCell : XlsBiffBlankCell
    {
        internal XlsBiffIntegerCell(byte[] bytes)
            : base(bytes)
        {
        }

        public override bool IsEmpty => false;

        /// <summary>
        /// Gets the cell value.
        /// </summary>
        public int Value => Id == BIFFRECORDTYPE.INTEGER_OLD ? ReadUInt16(0x7) : ReadUInt16(0x6);
    }
}