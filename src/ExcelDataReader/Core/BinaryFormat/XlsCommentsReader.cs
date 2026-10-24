using System.Text;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat;

internal static class XlsCommentsReader
{
    public static Dictionary<long, CellCommentThread> Load(XlsWorksheet worksheet)
    {
        var workbook = worksheet.Workbook;
        using var stream = XlsBiffStartup.Open(worksheet.Stream, out _, out _, (int)worksheet.DataOffset, workbook.BiffVersion, BIFFTYPE.Worksheet, secretKey: workbook.SecretKey, encryption: workbook.Encryption);
        var result = new Dictionary<long, CellCommentThread>();
        var notes = new Dictionary<int, (int Row, int Column, string Author)>();
        var texts = new Dictionary<int, string>();
        int objectId = -1;
        while (stream.ReadFrame(out var frame))
        {
            if (frame.Id == BIFFRECORDTYPE.EOF)
                break;
            switch (frame.Id)
            {
                case BIFFRECORDTYPE.NOTE:
                    if (workbook.BiffVersion < 8)
                    {
                        ReadLegacyNote(stream, frame, worksheet.Encoding, result);
                        break;
                    }

                    Require(frame.RecordSize >= 11);
                    int count = frame.ReadUInt16(8);
                    bool wide = (frame.ReadByte(10) & 1) != 0;
                    int size = count * (wide ? 2 : 1);
                    Require(size <= frame.RecordSize - 11);
                    var authorText = new StringBuilder(count);
                    for (int i = 0; i < count; i++)
                        authorText.Append(wide ? (char)frame.ReadUInt16(11 + (i * 2)) : (char)frame.ReadByte(11 + i));
                    string author = authorText.ToString();
                    int noteId = frame.ReadUInt16(6);
                    Require(!notes.ContainsKey(noteId));
                    notes.Add(noteId, (frame.ReadUInt16(0), frame.ReadUInt16(2), author));
                    break;
                case BIFFRECORDTYPE.OBJ when workbook.BiffVersion >= 8:
                    objectId = -1;
                    Require(frame.RecordSize >= 22 && frame.ReadUInt16(0) == 0x15 && frame.ReadUInt16(2) == 18);
                    if (frame.ReadUInt16(4) == 0x19)
                        objectId = frame.ReadUInt16(6);
                    break;
                case BIFFRECORDTYPE.TXO when objectId >= 0:
                    Require(!texts.ContainsKey(objectId));
                    texts.Add(objectId, ReadText(stream, frame));
                    objectId = -1;
                    break;
            }
        }

        foreach (var note in notes)
        {
            Require(texts.TryGetValue(note.Key, out string? text));
            CommentStore.Add(result, note.Value.Row, note.Value.Column, note.Value.Author, text!);
        }

        return result;
    }

    private static string ReadText(XlsBiffStream stream, XlsBiffFrame frame)
    {
        Require(frame.RecordSize >= 16);
        int remaining = frame.ReadUInt16(10);
        int runBytes = frame.ReadUInt16(12);
        Require(remaining == 0 ? runBytes == 0 : runBytes >= 16 && runBytes % 8 == 0);
        var text = new StringBuilder(remaining);
        while (remaining > 0)
        {
            Require(stream.ReadFrame(out frame) && frame.Id == BIFFRECORDTYPE.CONTINUE && frame.RecordSize > 1);
            bool wide = (frame.ReadByte(0) & 1) != 0;
            int bytes = frame.RecordSize - 1;
            Require(!wide || bytes % 2 == 0);
            int characters = bytes / (wide ? 2 : 1);
            Require(characters <= remaining);

            // Preserve UTF-16 code units even when a surrogate pair crosses records.
            for (int i = 0; i < characters; i++)
                text.Append(wide ? (char)frame.ReadUInt16(1 + (i * 2)) : (char)frame.ReadByte(i + 1));

            remaining -= characters;
        }

        while (runBytes > 0)
        {
            Require(stream.ReadFrame(out frame) && frame.Id == BIFFRECORDTYPE.CONTINUE && frame.RecordSize > 0 && frame.RecordSize <= runBytes);
            runBytes -= frame.RecordSize;
        }

        return text.ToString();
    }

    private static void ReadLegacyNote(XlsBiffStream stream, XlsBiffFrame frame, Encoding encoding, Dictionary<long, CellCommentThread> result)
    {
        Require(frame.RecordSize >= 6);
        int row = frame.ReadUInt16(0);
        Require(row != ushort.MaxValue);
        int column = frame.ReadUInt16(2);
        int remaining = frame.ReadUInt16(4);
        int length = Math.Min(remaining, 2048);
        using var bytes = new MemoryStream();
        while (true)
        {
            Require(frame.RecordSize == length + 6);
            bytes.Write(frame.Bytes, 10, length);
            remaining -= length;
            if (remaining == 0)
                break;
            Require(stream.ReadFrame(out frame) && frame.Id == BIFFRECORDTYPE.NOTE && frame.RecordSize > 6 && frame.ReadUInt16(0) == ushort.MaxValue);
            length = frame.ReadUInt16(4);
            Require(length is > 0 and <= 2048 && length <= remaining);
        }

        CommentStore.Add(result, row, column, null, encoding.GetString(bytes.ToArray()));
    }

    private static void Require(bool condition)
    {
        if (!condition)
            throw new ExcelReaderException("Invalid XLS comment record.");
    }
}