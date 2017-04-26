using ExcelDataReader.Portable.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReader.Portable
{
    public interface IExcelDataReader : IDataReader
    {
        /// <summary>
        /// Initializes the instance with specified file stream.
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        Task InitializeAsync(Stream fileStream);

        ///// <summary>
        ///// Read all data in to DataSet and return it
        ///// </summary>
        ///// <returns>The DataSet</returns>
        Task LoadDataSetAsync(IDatasetHelper datasetHelper);

        ///// <summary>
        /////Read all data in to DataSet and return it
        ///// </summary>
        ///// <param name="convertOADateTime">if set to <c>true</c> [try auto convert OA date time format].</param>
        ///// <returns>The DataSet</returns>
        Task LoadDataSetAsync(IDatasetHelper datasetHelper, bool convertOADateTime);

        /// <summary>
        /// Gets a value indicating whether file stream is valid.
        /// </summary>
        /// <value><c>true</c> if file stream is valid; otherwise, <c>false</c>.</value>
        bool IsValid { get; }

        /// <summary>
        /// Gets the exception message in case of error.
        /// </summary>
        /// <value>The exception message.</value>
        string ExceptionMessage { get; }

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

        /// <summary>
        /// Gets or sets a value indicating whether the first row contains the column names.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if the first row contains column names; otherwise, <c>false</c>.
        /// </value>
        bool IsFirstRowAsColumnNames { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether empty tables are allowed or skipped.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if empty tables are allowed; otherwise, <c>false</c>.
        /// </value>
        bool DoAllowEmptyTables { get; set; }

        /// <summary>
        /// Should OADates be converted to dates
        /// </summary>
        bool ConvertOaDate { get; set; }

        ReadOption ReadOption { get; set; }
        Encoding Encoding { get; }
        Encoding DefaultEncoding { get; }
    }
}