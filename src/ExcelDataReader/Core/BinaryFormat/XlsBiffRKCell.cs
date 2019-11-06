using System;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents an RK number cell
    /// </summary>
    internal class XlsBiffRKCell : XlsBiffBlankCell
    {
        internal XlsBiffRKCell(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the value of this cell
        /// </summary>
        public double Value => NumFromRK(ReadUInt32(0x6));

        /// <summary>
        /// Decodes RK-encoded number
        /// </summary>
        /// <param name="rk">Encoded number</param>
        /// <returns>The number.</returns>
        public static double NumFromRK(uint rk)
        {
            double num;
            if ((rk & 0x2) == 0x2)
            {
                num = (int)(rk >> 2 | ((rk & 0x80000000) == 0 ? 0 : 0xC0000000));
            }
            else
            {
                // hi words of IEEE num
                num = BitConverter.Int64BitsToDouble((long)(rk & 0xfffffffc) << 32);
            }

            if ((rk & 0x1) == 0x1)
                num /= 100; // divide by 100

            return num;
        }
    }
}