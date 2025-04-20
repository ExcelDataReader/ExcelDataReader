using ExcelDataReader.Core;

namespace ExcelDataReader;

/// <summary>
/// A range for cells using 0 index positions. 
/// </summary>
/// <param name="FromColumn">The column the range starts in.</param>
/// <param name="FromRow">The row the range starts in.</param>
/// <param name="ToColumn">The column the range ends in.</param>
/// <param name="ToRow">The row the range ends in.</param>
public sealed record CellRange(int FromColumn, int FromRow, int ToColumn, int ToRow)
{
    /// <inheritsdoc/>
    public override string ToString() => $"{FromRow}, {ToRow}, {FromColumn}, {ToColumn}";

#if NET8_0_OR_GREATER
    internal static CellRange Parse(string range)
    {
        int index = range.IndexOf(':');
        if (index >= 0 && range.IndexOf(':', index + 1) < 0)
        {
            ReadOnlySpan<char> span = range;
            ReferenceHelper.ParseReference(span[..index], out int fromColumn, out int fromRow);
            ReferenceHelper.ParseReference(span[(index + 1)..], out int toColumn, out int toRow);

            // 0 indexed vs 1 indexed
            return new(fromColumn - 1, fromRow - 1, toColumn - 1, toRow - 1);
        }

        return new(0, 0, 0, 0);
    }
#else
    internal static CellRange Parse(string range)
    {
        var fromTo = range.Split(':');
        if (fromTo.Length == 2)
        {
            ReferenceHelper.ParseReference(fromTo[0], out int fromColumn, out int fromRow);
            ReferenceHelper.ParseReference(fromTo[1], out int toColumn, out int toRow);

            // 0 indexed vs 1 indexed
            return new(fromColumn - 1, fromRow - 1, toColumn - 1, toRow - 1);
        }

        return new(0, 0, 0, 0);
    }
#endif
}