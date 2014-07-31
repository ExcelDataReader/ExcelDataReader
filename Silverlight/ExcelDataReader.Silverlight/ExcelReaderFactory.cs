namespace ExcelDataReader.Silverlight
{
	using System.IO;

	/// <summary>
	/// The ExcelReader Factory
	/// </summary>
	public static class ExcelReaderFactory
	{
		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public static IExcelDataReader CreateBinaryReader(Stream fileStream)
		{
			IExcelDataReader reader = new ExcelBinaryReader();
			reader.Initialize(fileStream);

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelOpenXmlReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public static IExcelDataReader CreateOpenXmlReader(Stream fileStream)
		{
			IExcelDataReader reader = new ExcelOpenXmlReader();
			reader.Initialize(fileStream);

			return reader;
		}
	}
}