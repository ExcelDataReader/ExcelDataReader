namespace ExcelDataReader.Core
{
    internal class Cell
    {
        /// <summary>
        /// Gets or sets the zero-based column index.
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// Gets or sets the index of the XF record describing the styling of this cell.
        /// </summary>
        public int XfIndex { get; set; }

        /// <summary>
        /// Gets or sets the effective style on the cell. The effective style is determined from 
        /// the Cell XF, with optional overrides from a Cell Style XF.
        /// </summary>
        public ExtendedFormat EffectiveStyle { get; set; }

        public object Value { get; set; }
    }
}
