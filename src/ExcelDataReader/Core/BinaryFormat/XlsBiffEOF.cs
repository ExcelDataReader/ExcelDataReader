namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents BIFF EOF resord.
    /// </summary>
    internal sealed class XlsBiffEof : XlsBiffRecord
    {
        internal XlsBiffEof(byte[] bytes)
            : base(bytes)
        {
        }
    }
}
