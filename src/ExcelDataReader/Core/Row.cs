using System;
using System.Collections.Generic;

namespace ExcelDataReader.Core
{
    internal class Row
    {
        /// <summary>
        /// Gets or sets the zero-based row index.
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// Gets or sets the height of this row in points. Zero if hidden or collapsed.
        /// </summary>
        public double Height { get; set; }

        /// <summary>
        /// Gets or sets the cells in this row.
        /// </summary>
        public List<Cell> Cells { get; set; }

        /// <summary>
        /// Gets a value indicating whether the row is empty. NOTE: Returns true if there are empty, but formatted cells.
        /// </summary>
        public bool IsEmpty
        {
            get
            {
                return Cells.Count == 0;
            }
        }

        /// <summary>
        /// Returns the zero-based maximum column index reference on this row.
        /// </summary>
        public int GetMaxColumnIndex()
        {
            int columnIndex = int.MinValue;
            foreach (var cell in Cells)
            {
                columnIndex = Math.Max(cell.ColumnIndex, columnIndex);
            }

            return columnIndex;
        }
    }
}
