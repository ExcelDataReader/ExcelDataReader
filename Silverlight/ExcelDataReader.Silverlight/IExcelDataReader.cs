namespace ExcelDataReader.Silverlight
{
	using System.IO;
	using Data;

	public interface IExcelDataReader //: IDataReader
	{
		/// <summary>
		/// Initializes the instance with specified file stream.
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		void Initialize(Stream fileStream);

		/// <summary>
		/// Initializes the instance with specified file stream.
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <param name="closeOnFail">If set to true, will close the file stream on fail. Otherwise, leaves it open.</param>
		void Initialize(Stream fileStream, bool closeOnFail);

		/// <summary>
		/// Read all data in to an IWorkBook and return it
		/// </summary>
		/// <returns>The DataSet</returns>
		IWorkBook AsWorkBook();

		/// <summary>
		/// Read all data in to an IWorkBook and return it
		/// </summary>
		/// <param name="convertOaDateTime">if set to <c>true</c> [try auto convert OA date time format].</param>
		/// <returns>The DataSet</returns>
		IWorkBook AsWorkBook(bool convertOaDateTime);

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
		/// Gets or sets an IWorkBookFactory instance for generating WorkBook and associated data classes for the AsWorkBook() method.
		/// </summary>
		IWorkBookFactory WorkBookFactory { get; set; }

		/// <summary>
		/// Closes this reader and its underlying stream.
		/// </summary>
		void Close();
	}
}