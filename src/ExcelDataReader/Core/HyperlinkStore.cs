using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core;

internal sealed class HyperlinkStore
{
    private readonly List<Entry> _entries = [];

    // Sorted row intervals form an implicit balanced tree, with a maximum end row per subtree.
    private int[]? _maxRows;

    public static string JoinAddress(string? address, string? location) =>
        string.IsNullOrEmpty(location) ? address ?? string.Empty : address + "#" + location;

    public static CellRange ParseRange(string reference)
    {
        int colon = reference.IndexOf(':');
        string start = colon < 0 ? reference : reference[..colon];
        string end = colon < 0 ? reference : reference[(colon + 1)..];
        if (!TryParseCell(start, out int fromColumn, out int fromRow) ||
            !TryParseCell(end, out int toColumn, out int toRow) ||
            fromColumn > toColumn || fromRow > toRow)
            throw new ExcelReaderException("Invalid hyperlink cell reference.");
        return new CellRange(fromColumn, fromRow, toColumn, toRow);
    }

    public void Add(CellRange range, CellHyperlink hyperlink)
    {
        if (_maxRows != null)
            throw new InvalidOperationException("Hyperlink metadata has already been indexed.");
        if (range.FromRow < 0 || range.FromColumn < 0 || range.ToRow < range.FromRow ||
            range.ToColumn < range.FromColumn || range.ToRow >= 1048576 || range.ToColumn >= 16384)
            throw new ExcelReaderException("Invalid hyperlink cell range.");
        _entries.Add(new Entry(range, hyperlink, _entries.Count));
    }

    public CellHyperlink? Find(int row, int column)
    {
        if (_entries.Count == 0)
            return null;
        if (_maxRows == null)
        {
            _entries.Sort((left, right) => left.Range.FromRow.CompareTo(right.Range.FromRow));
            _maxRows = new int[_entries.Count];
            Index(0, _entries.Count);
        }

        int ordinal = -1;
        CellHyperlink? result = null;
        Find(0, _entries.Count, row, column, ref ordinal, ref result);
        return result;
    }

    private static bool TryParseCell(string reference, out int column, out int row)
    {
        column = 0;
        row = 0;
        int position = 0;
        while (position < reference.Length && (uint)((reference[position] | 0x20) - 'a') < 26)
        {
            column = (column * 26) + (reference[position++] | 0x20) - 'a' + 1;
            if (column > 16384)
                return false;
        }

        if (column == 0 || position == reference.Length)
            return false;
        for (; position < reference.Length; position++)
        {
            int digit = reference[position] - '0';
            if ((uint)digit > 9)
                return false;
            row = (row * 10) + digit;
            if (row > 1048576)
                return false;
        }

        if (row == 0)
            return false;
        column--;
        row--;
        return true;
    }

    private int Index(int start, int end)
    {
        if (start == end)
            return -1;
        int middle = start + ((end - start) / 2);
        int maxRow = Math.Max(_entries[middle].Range.ToRow, Math.Max(Index(start, middle), Index(middle + 1, end)));
        _maxRows![middle] = maxRow;
        return maxRow;
    }

    private void Find(int start, int end, int row, int column, ref int ordinal, ref CellHyperlink? result)
    {
        if (start == end)
            return;
        int middle = start + ((end - start) / 2);
        if (_entries[start].Range.FromRow > row || _maxRows![middle] < row)
            return;
        var entry = _entries[middle];
        if (entry.Ordinal > ordinal && entry.Range.FromRow <= row && entry.Range.ToRow >= row &&
            entry.Range.FromColumn <= column && entry.Range.ToColumn >= column)
        {
            ordinal = entry.Ordinal;
            result = entry.Hyperlink;
        }

        Find(start, middle, row, column, ref ordinal, ref result);
        Find(middle + 1, end, row, column, ref ordinal, ref result);
    }

    private readonly record struct Entry(CellRange Range, CellHyperlink Hyperlink, int Ordinal);
}
