using System;

namespace ExcelDataReader.Core
{
    internal class Cell
    {
        public Cell(int columnIndex, int numberFormatIndex, object value)
        {
            ColumnIndex = columnIndex;
            NumberFormatIndex = numberFormatIndex;
            Value = value;
        }

        /// <summary>
        /// Gets the zero-based column index.
        /// </summary>
        public int ColumnIndex { get; }

        public int NumberFormatIndex { get; }

        public object Value { get; }
    }
}
