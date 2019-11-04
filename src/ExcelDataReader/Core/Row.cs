using System;
using System.Collections.Generic;

#nullable enable

namespace ExcelDataReader.Core
{
    internal class Row
    {
        public Row(int rowIndex, double height, List<Cell> cells) 
        {
            RowIndex = rowIndex;
            Height = height;
            Cells = cells;
        }

        /// <summary>
        /// Gets the zero-based row index.
        /// </summary>
        public int RowIndex { get; }

        /// <summary>
        /// Gets the height of this row in points. Zero if hidden or collapsed.
        /// </summary>
        public double Height { get; }

        /// <summary>
        /// Gets the cells in this row.
        /// </summary>
        public List<Cell> Cells { get; }
    }
}
