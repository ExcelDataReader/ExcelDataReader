using System;
using ExcelDataReader.Core;

namespace ExcelDataReader
{
    /// <summary>
    /// A range for cells using 0 index positions. 
    /// </summary>
    public sealed class CellRange
    {
        internal CellRange(string from, string to)
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

        internal CellRange(int fromColumn, int fromRow, int toColumn, int toRow)
        {
            FromColumn = fromColumn;
            FromRow = fromRow;
            ToColumn = toColumn;
            ToRow = toRow;
        }

        /// <summary>
        /// Gets the column the range starts in
        /// </summary>
        public int FromColumn { get; private set; }

        /// <summary>
        /// Gets the row the range starts in
        /// </summary>
        public int FromRow { get; private set; }

        /// <summary>
        /// Gets the column the range ends in
        /// </summary>
        public int ToColumn { get; private set; }

        /// <summary>
        /// Gets the row the range ends in
        /// </summary>
        public int ToRow { get; private set; }
    }
}