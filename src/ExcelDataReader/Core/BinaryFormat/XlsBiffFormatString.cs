using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents a string value of format.
/// </summary>
internal sealed class XlsBiffFormatString : XlsBiffRecord
{
    private readonly IXlsString _xlsString;

    internal XlsBiffFormatString(byte[] bytes, int biffVersion)
        : base(bytes)
    {
        if (Id == BIFFRECORDTYPE.FORMAT_V23)
        {
            // BIFF2-3
            _xlsString = new XlsShortByteString(bytes, ContentOffset);
        }
        else if (biffVersion >= 2 && biffVersion <= 5)
        {
            // BIFF4-5, or if there is a newer format record in a BIFF2-3 stream
            _xlsString = new XlsShortByteString(bytes, ContentOffset + 2);
        }
        else if (biffVersion == 8)
        {
            // BIFF8
            _xlsString = new XlsUnicodeString(bytes, ContentOffset + 2);
        }
        else
        {
            throw new ArgumentException("Unexpected BIFF version " + biffVersion, nameof(biffVersion));
        }
    }

    public ushort Index => Id switch
    {
        BIFFRECORDTYPE.FORMAT_V23 => throw new NotSupportedException("Index is not available for BIFF2 and BIFF3 FORMAT records."),
        _ => ReadUInt16(0),
    };

    /// <summary>
    /// Gets the string value.
    /// </summary>
    public string GetValue(Encoding encoding)
    {
        return _xlsString.GetValue(encoding);
    }

    #if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
    public override void Return()
    {        
    }
    #endif
}