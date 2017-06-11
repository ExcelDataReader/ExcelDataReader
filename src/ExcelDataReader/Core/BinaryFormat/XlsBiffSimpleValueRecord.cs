namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents record with the only two-bytes value
    /// </summary>
    internal class XlsBiffSimpleValueRecord : XlsBiffRecord
    {
        internal XlsBiffSimpleValueRecord(byte[] bytes, uint offset)
            : base(bytes, offset)
        {
        }

        /// <summary>
        /// Returns value
        /// </summary>
        public ushort Value
        {
            get { return ReadUInt16(0x0); }
        }
    }
}
