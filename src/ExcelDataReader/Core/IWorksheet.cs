namespace ExcelDataReader.Core
{
    /// <summary>
    /// The common worksheet interface between the binary and OpenXml formats.
    /// </summary>
    internal interface IWorksheet
    {
        string Name { get; }

        string CodeName { get; }

        string VisibleState { get; }

        HeaderFooter HeaderFooter { get; }

        int FieldCount { get; }

        int RowCount { get; }

        /// <summary>
        /// Gets the index of first row.
        /// </summary>
        int FirstRow { get; }

        /// <summary>
        /// Gets the index of last row + 1.
        /// </summary>
        int LastRow { get; }

        /// <summary>
        /// Gets the index of first column.
        /// </summary>
        int FirstColumn { get; }

        /// <summary>
        /// Gets the index of last column + 1.
        /// </summary>
        int LastColumn { get; }

        CellRange[] MergeCells { get; }

        Column[] ColumnWidths { get; }

        IEnumerable<Row> ReadRows();
    }
}
