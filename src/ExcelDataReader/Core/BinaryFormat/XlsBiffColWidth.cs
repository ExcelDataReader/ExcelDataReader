namespace ExcelDataReader.Core.BinaryFormat;

internal sealed class XlsBiffColWidth : XlsBiffRecord
{
    public XlsBiffColWidth(byte[] bytes)
        : base(bytes)
    {
        var colFirst = ReadByte(0x0);
        var colLast = ReadByte(0x1);
        var width = ReadUInt16(0x3);
        Value = new Column(colFirst, colLast, false, width);
    }

    public Column Value { get; }
}
