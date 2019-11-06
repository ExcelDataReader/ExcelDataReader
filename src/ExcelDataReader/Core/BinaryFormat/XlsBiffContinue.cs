namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents additional space for very large records
    /// </summary>
    internal class XlsBiffContinue : XlsBiffRecord
    {
        internal XlsBiffContinue(byte[] bytes)
            : base(bytes)
        {
        }
    }
}
