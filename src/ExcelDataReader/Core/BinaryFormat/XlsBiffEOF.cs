namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents BIFF EOF resord
    /// </summary>
    internal class XlsBiffEof : XlsBiffRecord
    {
        internal XlsBiffEof(byte[] bytes)
            : base(bytes)
        {
        }
    }
}
