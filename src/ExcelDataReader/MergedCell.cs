using System;
using ExcelDataReader.Core;

namespace ExcelDataReader
{
    /// <summary>
    /// Header and footer text. 
    /// </summary>
    public sealed class MergedCell
    {
        internal MergedCell(string from, string to)
        {
            int fromColumn, fromRow, toColumn, toRow;
            ReferenceHelper.ParseReference(from, out fromColumn, out fromRow);

            // 0 indexed vs 1 indexed
            FromColumn = fromColumn - 1;
            FromRow = fromRow - 1;

            ReferenceHelper.ParseReference(to, out toColumn, out toRow);

            // 0 indexed vs 1 indexed
            ToColumn = toColumn - 1;
            ToRow = toRow - 1;
        }

        internal MergedCell(int fromColumn, int fromRow, int toColumn, int toRow)
        {
            FromColumn = fromColumn;
            FromRow = fromRow;
            ToColumn = toColumn;
            ToRow = toRow;
        }

        public int FromColumn { get; private set; }

        public int FromRow { get; private set; }

        public int ToColumn { get; private set; }

        public int ToRow { get; private set; }

    }
}