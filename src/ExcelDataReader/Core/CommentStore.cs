using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core;

internal static class CommentStore
{
    public static void Add(Dictionary<long, CellCommentThread> store, int row, int column, string? author, string text)
    {
        if (row < 0 || column < 0 || row >= 1048576 || column >= 16384)
            throw new ExcelReaderException("Invalid comment cell reference.");
        long key = ((long)row << 32) | (uint)column;
        if (store.ContainsKey(key))
            throw new ExcelReaderException("Duplicate legacy comment cell reference.");
        store.Add(key, new CellCommentThread(author, text));
    }
}