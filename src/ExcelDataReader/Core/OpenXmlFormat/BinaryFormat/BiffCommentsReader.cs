using System.Text;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;

internal sealed class BiffCommentsReader(Stream stream) : BiffReader(stream)
{
    private enum State
    {
        Start,
        AuthorsStart,
        Authors,
        ListStart,
        List,
        Comment,
        End,
        Complete,
    }

    public Dictionary<long, CellCommentThread> Load()
    {
        var result = new Dictionary<long, CellCommentThread>();
        var authors = new List<string>();
        int row = -1;
        int column = -1;
        int author = -1;
        string text = string.Empty;
        bool hasText = false;
        State state = State.Start;
        while (TryReadFrame(out var buffer, out uint id, out uint length))
        {
            switch (id)
            {
                case 0x274:
                    Require(state == State.Start);
                    state = State.AuthorsStart;
                    break;
                case 0x276:
                    Require(state == State.AuthorsStart);
                    state = State.Authors;
                    break;
                case 0x278: // MS-XLSB BrtCommentAuthor
                    Require(state == State.Authors);
                    authors.Add(ReadText(buffer, length, 0));
                    break;
                case 0x277:
                    Require(state == State.Authors);
                    state = State.ListStart;
                    break;
                case 0x279:
                    Require(state == State.ListStart);
                    state = State.List;
                    break;
                case 0x27B: // BrtBeginComment
                    Require(length == 36 && state == State.List);
                    author = GetInt32(buffer, 0);
                    row = GetInt32(buffer, 4);
                    column = GetInt32(buffer, 12);
                    Require((uint)author < (uint)authors.Count && row == GetInt32(buffer, 8) && column == GetInt32(buffer, 16));
                    text = string.Empty;
                    hasText = false;
                    state = State.Comment;
                    break;
                case 0x27D: // BrtCommentText / RichStr
                    Require(state == State.Comment && !hasText && length >= 9 && (buffer[0] & 3) == 1);
                    text = ReadText(buffer, length, 1);
                    uint runsOffset = 5 + (GetDWord(buffer, 1) * 2);
                    Require(length >= runsOffset + 4 && (length - runsOffset - 4) % 4 == 0 && GetDWord(buffer, runsOffset) == (length - runsOffset - 4) / 4);
                    hasText = true;
                    break;
                case 0x27C: // BrtEndComment
                    Require(state == State.Comment);
                    CommentStore.Add(result, row, column, authors[author], text);
                    row = -1;
                    state = State.List;
                    break;
                case 0x27A:
                    Require(state == State.List);
                    state = State.End;
                    break;
                case 0x275:
                    Require(state == State.End);
                    state = State.Complete;
                    break;
            }
        }

        Require(state == State.Complete);
        return result;
    }

    private static string ReadText(byte[] buffer, uint length, uint offset)
    {
        Require(length >= offset + 4);
        uint characters = GetDWord(buffer, offset);
        Require(characters <= (length - offset - 4) / 2);
        return Helpers.ConvertEscapeChars(Encoding.Unicode.GetString(buffer, (int)offset + 4, (int)characters * 2));
    }

    private static void Require(bool condition)
    {
        if (!condition)
            throw new ExcelReaderException("Invalid binary comment record.");
    }
}