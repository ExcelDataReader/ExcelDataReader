namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents InterfaceHdr record in Wokrbook Globals
    /// </summary>
    internal class XlsBiffInterfaceHdr : XlsBiffRecord
    {
        internal XlsBiffInterfaceHdr(byte[] bytes, uint offset)
            : base(bytes, offset)
        {
        }

        /// <summary>
        /// Gets the CodePage for Interface Header
        /// </summary>
        public ushort CodePage => ReadUInt16(0x0);
    }
}
