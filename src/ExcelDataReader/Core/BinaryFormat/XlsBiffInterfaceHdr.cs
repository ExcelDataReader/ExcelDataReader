namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents InterfaceHdr record in Wokrbook Globals
    /// </summary>
    internal class XlsBiffInterfaceHdr : XlsBiffRecord
    {
        internal XlsBiffInterfaceHdr(byte[] bytes)
            : base(bytes)
        {
        }

        /// <summary>
        /// Gets the CodePage for Interface Header
        /// </summary>
        public ushort CodePage => ReadUInt16(0x0);
    }
}
