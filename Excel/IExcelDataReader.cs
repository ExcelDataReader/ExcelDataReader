using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;

using Excel.Core;

namespace Excel
{
    public interface IExcelDataReader : IDataReader
    {
        #region Properties for batch support
        int BatchSize { get; set; }
        int SheetIndex { get; set; }
        string SheetName { get; set; }
        int SkipRows { get; set; }
        #endregion

        /// <summary>
        /// Initializes the instance with specified file stream.
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        void Initialize(Stream fileStream);

        /// <summary>
        /// Read all data in to DataSet and return it
        /// </summary>
        /// <returns>The DataSet</returns>
        DataSet AsDataSet();

        /// <summary>
        ///Read all data in to DataSet and return it
        /// </summary>
        /// <param name="convertOADateTime">if set to <c>true</c> [try auto convert OA date time format].</param>
        /// <returns>The DataSet</returns>
        DataSet AsDataSet(bool convertOADateTime);

        /// <summary>
        /// Gets a value indicating whether file stream is valid.
        /// </summary>
        /// <value><c>true</c> if file stream is valid; otherwise, <c>false</c>.</value>
        bool IsValid { get;}

        /// <summary>
        /// Gets the exception message in case of error.
        /// </summary>
        /// <value>The exception message.</value>
        string ExceptionMessage { get;}

        /// <summary>
        /// Gets the sheet name.
        /// </summary>
        /// <value>The sheet name.</value>
        string Name { get;}

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
        int ResultsCount { get;}

        /// <summary>
        /// Gets or sets a value indicating whether the first row contains the column names.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if the first row contains column names; otherwise, <c>false</c>.
        /// </value>
        bool IsFirstRowAsColumnNames { get;set;}

        #region Methods for batch support
        DataSet GetSchema(bool isFirstRowAsColumnNames = true, int skipRows = 0);
        DataSet GetSchema(List<SheetParameters> sheetParametersList);
        DataTable GetSchema(SheetParameters sheetParameters);
        bool ReadBatch();
        DataTable GetCurrentBatch();

        DataTable GetTopRows(int rowCount, SheetParameters sheetParameters);

        List<string> GetSheetNames();

        #endregion
    }
}