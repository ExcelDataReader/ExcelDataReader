using System.Collections.Generic;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core
{
    /// <summary>
    /// The common workbook interface between the binary and OpenXml formats
    /// </summary>
    /// <typeparam name="TWorksheet">A type implementing IWorksheet</typeparam>
    internal interface IWorkbook<TWorksheet>
        where TWorksheet : IWorksheet
    {
        int ResultsCount { get; }

        IEnumerable<TWorksheet> ReadWorksheets();

        NumberFormatString GetNumberFormatString(int index);
    }
}
