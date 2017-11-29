using System.Collections.Generic;
using ExcelDataReader.Core.NumberFormat;

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

        List<MergedCell> MergedCells { get; }

        IEnumerable<Row> ReadRows();

        NumberFormatString GetNumberFormatString(int index);
    }
}
