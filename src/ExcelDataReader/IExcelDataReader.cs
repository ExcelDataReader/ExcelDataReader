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
        /// <value>The sheet name.</value>
        string Name { get; }

        /// <summary>
        /// Gets the state of the visible.
        /// </summary>
        /// <value>
        /// The state of the visible.
        /// </value>
        string VisibleState { get; }

        /// <summary>
        /// Gets the number of results (workbooks).
        /// </summary>
        /// <value>The results count.</value>
        int ResultsCount { get; }

        Encoding Encoding { get; }

        /// <summary>
        /// Seeks to the first result.
        /// </summary>
        void Reset();
    }
}