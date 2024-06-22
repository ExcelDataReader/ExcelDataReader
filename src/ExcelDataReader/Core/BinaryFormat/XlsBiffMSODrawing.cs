namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents MSO Drawing record.
    /// </summary>
    internal sealed class XlsBiffMSODrawing : XlsBiffRecord
    {
        internal XlsBiffMSODrawing(byte[] bytes)
            : base(bytes)
        {
        }
    }
}
