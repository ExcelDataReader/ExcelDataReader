namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents MSO Drawing record
    /// </summary>
    internal class XlsBiffMSODrawing : XlsBiffRecord
    {
        internal XlsBiffMSODrawing(byte[] bytes, uint offset)
            : base(bytes, offset)
        {
        }
    }
}
