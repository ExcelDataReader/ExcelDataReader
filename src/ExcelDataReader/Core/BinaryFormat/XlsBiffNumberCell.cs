namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a floating-point number 
    /// </summary>
    internal class XlsBiffNumberCell : XlsBiffBlankCell
    {
        internal XlsBiffNumberCell(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the value of this cell
        /// </summary>
        public double Value => Id == BIFFRECORDTYPE.NUMBER_OLD ? ReadDouble(0x7) : ReadDouble(0x6);
    }
}