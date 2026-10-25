using System.Text;
using ExcelDataReader.Core;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;

internal sealed class BiffHyperlinksReader(Stream stream, IReadOnlyDictionary<string, string> relationships) : BiffReader(stream)
{
    private const uint BrtHLink = 494;

    public HyperlinkStore Load()
    {
        var result = new HyperlinkStore();
        while (TryReadFrame(out var buffer, out uint id, out uint length))
        {
            if (id != BrtHLink)
                continue;

            uint offset = 0;
            Require(length >= 16);
            uint fromRow = GetDWord(buffer, offset);
            uint toRow = GetDWord(buffer, offset + 4);
            uint fromColumn = GetDWord(buffer, offset + 8);
            uint toColumn = GetDWord(buffer, offset + 12);
            offset += 16;
            Require(fromRow <= toRow && toRow < 1048576 &&
                fromColumn <= toColumn && toColumn < 16384);
            var range = new CellRange((int)fromColumn, (int)fromRow, (int)toColumn, (int)toRow);

            string? relationshipId = ReadNullableWideString(buffer, length, ref offset);
            string location = ReadWideString(buffer, length, ref offset);
            string tooltip = ReadWideString(buffer, length, ref offset);
            _ = ReadWideString(buffer, length, ref offset);
            Require(offset == length);

            string? address;
            if (relationshipId is { Length: > 0 })
            {
                if (!relationships.TryGetValue(relationshipId, out var target))
                    continue;
                address = HyperlinkStore.JoinAddress(target, location);
            }
            else
            {
                address = HyperlinkStore.JoinAddress(null, location);
                if (location.Length == 0 && tooltip.Length == 0)
                    continue;
            }

            result.Add(range, new CellHyperlink(address, tooltip.Length == 0 ? null : tooltip));
        }

        return result;
    }

    private static string? ReadNullableWideString(byte[] buffer, uint length, ref uint offset)
    {
        RequireRemaining(length, offset, 4);
        uint characters = GetDWord(buffer, offset);
        offset += 4;
        if (characters == uint.MaxValue)
            return null;
        Require(characters <= 255);
        string value = ReadCharacters(buffer, length, ref offset, characters);
        Require(value.IndexOf('\0') < 0);
        return value;
    }

    private static string ReadWideString(byte[] buffer, uint length, ref uint offset)
    {
        RequireRemaining(length, offset, 4);
        uint characters = GetDWord(buffer, offset);
        offset += 4;
        Require(characters != uint.MaxValue);
        return ReadCharacters(buffer, length, ref offset, characters);
    }

    private static string ReadCharacters(byte[] buffer, uint length, ref uint offset, uint characters)
    {
        Require(characters <= (length - offset) / 2);
        int byteLength = checked((int)(characters * 2));
        string value = Encoding.Unicode.GetString(buffer, checked((int)offset), byteLength);
        offset += (uint)byteLength;
        return value;
    }

    private static void RequireRemaining(uint length, uint offset, uint required) =>
        Require(offset <= length && required <= length - offset);

    private static void Require(bool condition)
    {
        if (!condition)
            throw new ExcelReaderException("Invalid binary hyperlink record.");
    }
}
