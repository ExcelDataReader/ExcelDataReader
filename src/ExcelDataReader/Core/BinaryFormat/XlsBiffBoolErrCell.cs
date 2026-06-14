namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents a BOOLERR/BOOLERR_OLD cell record.
/// </summary>
internal sealed class XlsBiffBoolErrCell : XlsBiffBlankCell
{
    internal XlsBiffBoolErrCell(byte[] bytes)
        : base(bytes)
    {
    }

    public override bool IsEmpty => false;
}
