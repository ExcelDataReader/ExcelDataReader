using System;
using System.Data;
using System.Text;

namespace ExcelDataReader
{
    public interface IExcelDataReader : IDataReader
    {
        /// <summary>
        /// Gets the sheet name.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Gets the sheet visible state.
        /// </summary>
        string VisibleState { get; }

        /// <summary>
        /// Gets the sheet header and footer -or- <see langword="null"/> if none set.
        /// </summary>
        HeaderFooter HeaderFooter { get; }

        /// <summary>
        /// Gets the number of results (workbooks).
        /// </summary>
        int ResultsCount { get; }

        /// <summary>
        /// Seeks to the first result.
        /// </summary>
        void Reset();
    }
}