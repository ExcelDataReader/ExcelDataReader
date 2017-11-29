using System;
using ExcelDataReader.Core;

namespace ExcelDataReader
{
    /// <summary>
    /// Header and footer text. 
    /// </summary>
    internal sealed class MergedCell
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

        private Cell SourceCell { get; set; }

        /// <summary>
        /// If this cell column/row index is part of a merged a cell, get the cell if not the top left (contains the data)
        /// </summary>
        /// <param name="col">Cell column index</param>
        /// <param name="row">Cell row index</param>
        /// <param name="cell">Cell </param>
        /// <param name="merged">The cloned source cell</param>
        /// <returns>Whether this is a merged cell</returns>
        public bool GetSourceValue(int col, int row, Cell cell, out Cell merged)
        {
            merged = null;
            if (col == FromColumn && row == FromRow)
            {
                SourceCell = cell;
                return false;
            }

            if (SourceCell != null && col >= FromColumn && col <= ToColumn 
                && row >= FromRow && row <= ToRow)
            {
                merged = new Cell()
                {
                    ColumnIndex = col,
                    NumberFormatIndex = SourceCell.NumberFormatIndex,
                    Value = SourceCell.Value
                };
                
                return true;
            }

            return false;
        }
    }
}