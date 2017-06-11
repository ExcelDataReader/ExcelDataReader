namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// For now QuickTip will do nothing, it seems to have a different
    /// </summary>
    internal class XlsBiffQuickTip : XlsBiffRecord
    {
        internal XlsBiffQuickTip(byte[] bytes, uint offset)
            : base(bytes, offset)
        {
        }
    }
}