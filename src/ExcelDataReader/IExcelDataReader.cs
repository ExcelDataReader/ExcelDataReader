using System.IO;
using System.Text;
using ExcelDataReader.Data;

namespace Excel
{
	public interface IExcelDataReader : IDataReader
	{
		/// <summary>
		/// Initializes the instance with specified file stream.
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		void Initialize(Stream fileStream);

		/// <summary>
		/// Seeks to the first result.
		/// </summary>
		void Reset();

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

        /// <summary>
        /// Should OADates be converted to dates
        /// </summary>
        bool ConvertOaDate { get; set; }

        ReadOption ReadOption { get; set;  }
	    Encoding Encoding { get; }
	    Encoding DefaultEncoding { get; }
	}
}