namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents a floating-point number. 
/// </summary>
internal sealed class XlsBiffNumberCell : XlsBiffBlankCell
{
    internal XlsBiffNumberCell(byte[] bytes)
        : base(bytes)
    {
    }

    public override bool IsEmpty => false;

    /// <summary>
    /// Gets the value of this cell.
    /// </summary>
    public double Value => new XlsCellFrame(Bytes).NumberValue;
}