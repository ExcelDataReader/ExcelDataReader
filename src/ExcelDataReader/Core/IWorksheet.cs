using System.Collections.Generic;

namespace ExcelDataReader.Core
{
    /// <summary>
    /// The common worksheet interface between the binary and OpenXml formats
    /// </summary>
    internal interface IWorksheet
    {
        string Name { get; }

        string CodeName { get; }

        string VisibleState { get; }

        HeaderFooter HeaderFooter { get; }

        int FieldCount { get; }

        int RowCount { get; }

        CellRange[] MergeCells { get; }

        Column[] ColumnWidths { get; }

        IEnumerable<Row> ReadRows();
    }
}
