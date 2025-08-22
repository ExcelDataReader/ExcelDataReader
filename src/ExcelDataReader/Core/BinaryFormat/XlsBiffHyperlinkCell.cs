using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents an Excel hyperlink cell, parsing the URL.
/// Supported for BIFF 8 only
/// </summary>
internal sealed class XlsBiffHyperlinkCell : XlsBiffRecord
{
    private readonly string _url;

    internal XlsBiffHyperlinkCell(byte[] bytes)
        : base(bytes)
    {
        // Row 4 bytes
        // Column 4 bytes
        // Hyperlink GUID 16 bytes
        // Flags 4 bytes
        // Display text block length 4 bytes
        // Total 32 bytes
        var offset = 32;

        // Display text length 4 bytes
        var textLength = ReadInt32(offset);
        offset += 4;
        
        // 2 bytes per character
        offset += textLength * 2;

        // Skip URL moniker GUID (16 bytes)
        offset += 16;

        // URL length 4 bytes
        int urlLength = ReadInt32(offset);
        offset += 4;

        byte[] urlBytes = ReadArray(offset, urlLength);

        // Decode UTF-16LE string
        _url = Encoding.Unicode.GetString(urlBytes).Split('\0').First();
    }

    public int Row => ReadUInt16(0x0);
    
    public int Col => ReadUInt16(0x4);
    
    public string GetUrl() => _url;
}