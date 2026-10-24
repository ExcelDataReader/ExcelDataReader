using System.Text;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat;

internal static class XlsFormulaStringReader
{
    public static string? Read(XlsBiffStream stream, int version, Encoding encoding)
    {
        if (!stream.ReadFrame(out var frame))
            return null;
        if (frame.Id == BIFFRECORDTYPE.SHAREDFMLA && !stream.ReadFrame(out frame))
            return null;
        if (frame.Id is not BIFFRECORDTYPE.STRING and not BIFFRECORDTYPE.STRING_OLD)
            return null;

        if (version is >= 2 and <= 5)
            return ReadLegacyString(stream, frame, version, encoding);
        if (version != 8)
            throw new ArgumentException("Unexpected BIFF version " + version, nameof(version));

        int characters = frame.ReadUInt16(0);
        if (characters == 0)
            return string.Empty;
        int byteCount = (frame.ReadStringByteAt(6) & 1) != 0 ? characters * 2 : characters;
        if (frame.RecordSize >= 3 + byteCount)
            return XlsStringData.Read(frame, 4, false, true).GetValue(encoding);

        var reader = new XlsSSTReader();
        var strings = new SharedStringTable();
        reader.ReadStrings(frame, 4, false, strings, 1);
        while (strings.Count == 0)
        {
            int position = stream.Position;
            if (!stream.ReadFrame(out frame))
                break;
            if (frame.Id != BIFFRECORDTYPE.CONTINUE)
            {
                stream.Position = position;
                break;
            }

            reader.ReadStrings(frame, 4, true, strings, 1);
        }

        reader.Flush(strings);
        if (strings.Count == 0)
            throw new ExcelReaderException(Errors.ErrorBiffStringSize);
        return strings[0];
    }

    private static string ReadLegacyString(XlsBiffStream stream, XlsBiffFrame frame, int version, Encoding encoding)
    {
        int characters = version == 2 ? frame.ReadStringByteAt(4) : frame.ReadUInt16(0);
        int offset = version == 2 ? 5 : 6;
        int byteCount = characters * (Helpers.IsSingleByteEncoding(encoding) ? 1 : 2);
        int available = frame.RecordSize + 4 - offset;
        if (available >= byteCount)
            return encoding.GetString(frame.Bytes, offset, byteCount);

        var bytes = new byte[byteCount];
        Buffer.BlockCopy(frame.Bytes, offset, bytes, 0, available);
        int copied = available;
        while (copied < byteCount)
        {
            int position = stream.Position;
            if (!stream.ReadFrame(out frame))
                break;
            if (frame.Id != BIFFRECORDTYPE.CONTINUE)
            {
                stream.Position = position;
                break;
            }

            int count = Math.Min(frame.RecordSize, byteCount - copied);
            Buffer.BlockCopy(frame.Bytes, 4, bytes, copied, count);
            copied += count;
        }

        if (copied != byteCount)
            throw new ExcelReaderException(Errors.ErrorBiffRecordSize);
        return encoding.GetString(bytes, 0, byteCount);
    }
}