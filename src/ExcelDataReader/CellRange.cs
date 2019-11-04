using System;
using ExcelDataReader.Core;

namespace ExcelDataReader
{
    /// <summary>
    /// A range for cells using 0 index positions. 
    /// </summary>
    public sealed class CellRange
    {
        internal CellRange(string range)
        {
            var fromTo = range.Split(':');
            if (fromTo.Length == 2)
            {
                ReferenceHelper.ParseReference(fromTo[0], out int column, out int row);
                
                // 0 indexed vs 1 indexed
                FromColumn = column - 1;
                FromRow = row - 1;

                ReferenceHelper.ParseReference(fromTo[1], out column, out row);

                // 0 indexed vs 1 indexed
                ToColumn = column - 1;
                ToRow = row - 1;
            }
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
        public int FromColumn { get; }

        /// <summary>
        /// Gets the row the range starts in
        /// </summary>
        public int FromRow { get; }

        /// <summary>
        /// Gets the column the range ends in
        /// </summary>
        public int ToColumn { get; }

        /// <summary>
        /// Gets the row the range ends in
        /// </summary>
        public int ToRow { get; }

        /// <inheritsdoc/>
        public override string ToString() => $"{FromRow}, {ToRow}, {FromColumn}, {ToColumn}";
    }
}