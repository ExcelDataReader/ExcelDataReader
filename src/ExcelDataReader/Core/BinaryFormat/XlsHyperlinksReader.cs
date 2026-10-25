using System.Text;
using ExcelDataReader.Core;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat;

internal static class XlsHyperlinksReader
{
    private static readonly Guid HyperlinkClassId = new("79EAC9D0-BAF9-11CE-8C82-00AA004BA90B");
    private static readonly Guid UrlMonikerClassId = new("79EAC9E0-BAF9-11CE-8C82-00AA004BA90B");
    private static readonly Guid FileMonikerClassId = new("00000303-0000-0000-C000-000000000046");

    public static HyperlinkStore Load(XlsWorksheet worksheet)
    {
        var workbook = worksheet.Workbook;
        if (workbook.BiffVersion < 8)
            return new HyperlinkStore();

        using var stream = XlsBiffStartup.Open(
            worksheet.Stream,
            out _,
            out _,
            (int)worksheet.DataOffset,
            workbook.BiffVersion,
            BIFFTYPE.Worksheet,
            secretKey: workbook.SecretKey,
            encryption: workbook.Encryption);

        var hyperlinks = new List<(CellRange Range, string Address)>();
        var tooltips = new Dictionary<CellRange, string>();
        while (stream.ReadFrame(out var frame))
        {
            if (frame.Id == BIFFRECORDTYPE.EOF)
                break;

            switch (frame.Id)
            {
                case BIFFRECORDTYPE.HLINK:
                    if (TryReadHyperlink(frame, GetAnsiPathEncoding(worksheet.Encoding), out var range, out var address))
                        hyperlinks.Add((range, address));
                    break;
                case BIFFRECORDTYPE.QUICKTIP:
                    var tooltip = ReadTooltip(frame, out var tooltipRange);
                    tooltips[tooltipRange] = tooltip;
                    break;
            }
        }

        var result = new HyperlinkStore();
        foreach (var hyperlink in hyperlinks)
        {
            tooltips.TryGetValue(hyperlink.Range, out string? tooltip);
            result.Add(
                hyperlink.Range,
                new CellHyperlink(hyperlink.Address, string.IsNullOrEmpty(tooltip) ? null : tooltip));
        }

        return result;
    }

    private static Encoding GetAnsiPathEncoding(Encoding worksheetEncoding)
    {
        // BIFF8 defaults to UTF-16 for BIFF strings, but FileMoniker ansiPath is byte-oriented.
        return worksheetEncoding.CodePage is 1200 or 1201 ? Encoding.GetEncoding(1252) : worksheetEncoding;
    }

    private static bool TryReadHyperlink(XlsBiffFrame frame, Encoding encoding, out CellRange range, out string address)
    {
        range = ReadRange(frame, 0, 8);
        address = string.Empty;
        if (frame.RecordSize < 32)
            throw InvalidRecord();

        if (ReadGuid(frame, 8) != HyperlinkClassId)
            return false;

        int offset = 24;
        int end = frame.RecordSize;
        Require(ReadUInt32(frame, ref offset, end) == 2);
        uint flags = ReadUInt32(frame, ref offset, end);
        bool hasMoniker = (flags & 1) != 0;
        bool hasLocation = (flags & (1u << 3)) != 0;
        bool monikerSavedAsString = (flags & (1u << 8)) != 0;

        if ((flags & (1u << 4)) != 0)
            ReadHyperlinkString(frame, ref offset, end);
        if ((flags & (1u << 7)) != 0)
            ReadHyperlinkString(frame, ref offset, end);

        string? destination = null;
        if (hasMoniker)
        {
            if (monikerSavedAsString)
            {
                destination = ReadHyperlinkString(frame, ref offset, end);
            }
            else
            {
                RequireRange(frame, offset, 16, end);
                var monikerClassId = ReadGuid(frame, offset);
                offset += 16;
                if (monikerClassId == UrlMonikerClassId)
                    destination = ReadUrlMoniker(frame, ref offset, end);
                else if (monikerClassId == FileMonikerClassId)
                    destination = ReadFileMoniker(frame, ref offset, end, encoding);
                else
                    return false;
            }
        }

        string? location = hasLocation ? ReadHyperlinkString(frame, ref offset, end) : null;
        if ((flags & (1u << 5)) != 0)
            Skip(frame, ref offset, 16, end);
        if ((flags & (1u << 6)) != 0)
            Skip(frame, ref offset, 8, end);

        address = HyperlinkStore.JoinAddress(destination, location);
        return true;
    }

    private static string ReadTooltip(XlsBiffFrame frame, out CellRange range)
    {
        if (frame.RecordSize < 12 || frame.ReadUInt16(0) != (ushort)BIFFRECORDTYPE.QUICKTIP)
            throw InvalidRecord();
        range = ReadRange(frame, 2, 10);
        int stringBytes = frame.RecordSize - 10;
        Require((stringBytes & 1) == 0 && stringBytes <= 512);
        Require(frame.ReadUInt16At(frame.RecordSize + 2) == 0);
        return Encoding.Unicode.GetString(frame.Bytes, 14, stringBytes - 2);
    }

    private static CellRange ReadRange(XlsBiffFrame frame, int offset, int minimumSize)
    {
        if (frame.RecordSize < minimumSize)
            throw InvalidRecord();
        int firstRow = frame.ReadUInt16(offset);
        int lastRow = frame.ReadUInt16(offset + 2);
        int firstColumn = frame.ReadUInt16(offset + 4);
        int lastColumn = frame.ReadUInt16(offset + 6);
        Require(firstColumn <= 255 && lastColumn <= 255 &&
            firstRow <= lastRow && firstColumn <= lastColumn);
        return new CellRange(firstColumn, firstRow, lastColumn, lastRow);
    }

    private static string ReadHyperlinkString(XlsBiffFrame frame, ref int offset, int end)
    {
        uint characterCount = ReadUInt32(frame, ref offset, end);
        Require(characterCount > 0 && characterCount <= (uint)((end - offset) / 2));
        int byteCount = (int)characterCount * 2;
        RequireRange(frame, offset, byteCount, end);
        int contentBytes = byteCount - 2;
        Require(frame.ReadUInt16At(4 + offset + contentBytes) == 0);
        string value = Encoding.Unicode.GetString(frame.Bytes, 4 + offset, contentBytes);
        offset += byteCount;
        return value;
    }

    private static string ReadUrlMoniker(XlsBiffFrame frame, ref int offset, int end)
    {
        uint length = ReadUInt32(frame, ref offset, end);
        Require(length is >= 2 and <= int.MaxValue);
        int monikerLength = (int)length;
        Require(monikerLength <= end - offset);
        Require((monikerLength & 1) == 0);
        int urlLength = 2;
        while (urlLength <= monikerLength && frame.ReadUInt16(offset + urlLength - 2) != 0)
            urlLength += 2;
        Require(urlLength <= monikerLength &&
            (urlLength == monikerLength || monikerLength - urlLength == 24));
        string url = Encoding.Unicode.GetString(frame.Bytes, 4 + offset, urlLength - 2);
        offset += monikerLength;
        return url;
    }

    private static string ReadFileMoniker(XlsBiffFrame frame, ref int offset, int end, Encoding encoding)
    {
        ushort antiCount = ReadUInt16(frame, ref offset, end);
        uint ansiLength = ReadUInt32(frame, ref offset, end);
        Require(ansiLength is > 0 and <= 32767 and <= int.MaxValue);
        int ansiByteCount = (int)ansiLength;
        RequireRange(frame, offset, ansiByteCount, end);
        Require(frame.Bytes[4 + offset + ansiByteCount - 1] == 0);
        for (int i = 0; i < ansiByteCount - 1; i++)
            Require(frame.Bytes[4 + offset + i] != 0);
        string ansiPath = encoding.GetString(frame.Bytes, 4 + offset, ansiByteCount - 1);
        offset += ansiByteCount;

        ushort endServer = ReadUInt16(frame, ref offset, end);
        Require(ReadUInt16(frame, ref offset, end) == 0xDEAD);
        Skip(frame, ref offset, 16, end);
        Skip(frame, ref offset, 4, end);
        uint unicodePathSize = ReadUInt32(frame, ref offset, end);
        Require(unicodePathSize <= int.MaxValue && unicodePathSize <= end - offset);

        if (unicodePathSize == 0)
            return AddParentDirectories(ansiPath, antiCount, endServer);

        Require(unicodePathSize >= 6);
        int unicodeEnd = offset + (int)unicodePathSize;
        uint unicodeByteCount = ReadUInt32(frame, ref offset, unicodeEnd);
        Require(unicodeByteCount <= int.MaxValue && unicodeByteCount <= unicodeEnd - offset - 2);
        Require(ReadUInt16(frame, ref offset, unicodeEnd) == 3);
        Require((unicodeByteCount & 1) == 0 && unicodeByteCount == unicodeEnd - offset);
        string unicodePath = Encoding.Unicode.GetString(frame.Bytes, 4 + offset, (int)unicodeByteCount);
        offset = unicodeEnd;
        return unicodePath;
    }

    private static string AddParentDirectories(string path, ushort antiCount, ushort endServer)
    {
        bool driveRooted = path.Length >= 3 && char.IsLetter(path[0]) && path[1] == ':' &&
            path[2] is '\\' or '/';
        if (antiCount == 0 || endServer != ushort.MaxValue || driveRooted ||
            (path.Length > 0 && path[0] is '\\' or '/'))
            return path;

        int existingParents = 0;
        int offset = 0;
        while (path.Length - offset >= 3 &&
            (string.CompareOrdinal(path, offset, @"..\", 0, 3) == 0 ||
             string.CompareOrdinal(path, offset, "../", 0, 3) == 0))
        {
            existingParents++;
            offset += 3;
        }

        int missingParents = Math.Max(0, antiCount - existingParents);
        if (missingParents == 0)
            return path;

        var prefix = new StringBuilder(missingParents * 3);
        for (int i = 0; i < missingParents; i++)
            prefix.Append(@"..\");
        return prefix.Append(path).ToString();
    }

    private static Guid ReadGuid(XlsBiffFrame frame, int offset)
    {
        RequireRange(frame, offset, 16, frame.RecordSize);
        var bytes = new byte[16];
        Buffer.BlockCopy(frame.Bytes, 4 + offset, bytes, 0, bytes.Length);
        return new Guid(bytes);
    }

    private static ushort ReadUInt16(XlsBiffFrame frame, ref int offset, int end)
    {
        RequireRange(frame, offset, 2, end);
        ushort value = frame.ReadUInt16(offset);
        offset += 2;
        return value;
    }

    private static uint ReadUInt32(XlsBiffFrame frame, ref int offset, int end)
    {
        RequireRange(frame, offset, 4, end);
        uint value = frame.ReadUInt32(offset);
        offset += 4;
        return value;
    }

    private static void Skip(XlsBiffFrame frame, ref int offset, int count, int end)
    {
        RequireRange(frame, offset, count, end);
        offset += count;
    }

    private static void RequireRange(XlsBiffFrame frame, int offset, int count, int end)
    {
        Require(offset >= 0 && count >= 0 && offset <= end - count && end <= frame.RecordSize);
    }

    private static void Require(bool condition)
    {
        if (!condition)
            throw InvalidRecord();
    }

    private static ExcelReaderException InvalidRecord() => new("Invalid XLS hyperlink record.");
}
