namespace ExcelDataReader.Core
{
    internal class Cell
    {
        /// <summary>
        /// Gets or sets the zero-based column index.
        /// </summary>
        public int ColumnIndex { get; set; }

        public int NumberFormatIndex { get; set; }

        public object Value { get; set; }
    }
}
