namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents record with the only two-bytes value
    /// </summary>
    internal class XlsBiffSimpleValueRecord : XlsBiffRecord
    {
        internal XlsBiffSimpleValueRecord(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the value
        /// </summary>
        public ushort Value => ReadUInt16(0x0);
    }
}
