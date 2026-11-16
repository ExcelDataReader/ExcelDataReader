using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

internal sealed class XlsFontInfo
{
    private ushort _codePage = 1252;
    private Encoding? _byteStringEncoding;

    public XlsFontInfo(XlsBiffFrame frame, int biffVersion)
    {
        if (biffVersion < 5)
            return;

        _codePage = frame.ReadByte(12) switch
        {
            0 or 1 => 1252,
            77 => 10000,
            128 => 932,
            129 => 949,
            130 => 1361,
            134 => 936,
            136 => 950,
            161 => 1253,
            162 => 1254,
            163 => 1258,
            177 => 1255,
            178 => 1256,
            186 => 1257,
            204 => 1251,
            222 => 874,
            238 => 1250,
            255 => 850,
            _ => 1252,
        };
    }

    public Encoding ByteStringEncoding => _byteStringEncoding ??= EncodingHelper.GetEncoding(_codePage);
}