using System.IO;
using ExcelDataReader.Portable.Data;
using ExcelDataReader.Portable.IO;
using PCLStorage;


namespace ExcelDataReader.Portable
{
	/// <summary>
	/// The ExcelReader Factory
	/// </summary>
	public class ExcelReaderFactory
	{
	    private readonly IDataHelper dataHelper;
	    private readonly IFileHelper fileHelper;
	    private readonly IFileSystem fileSystem;

	    public ExcelReaderFactory(IDataHelper dataHelper, IFileHelper fileHelper, IFileSystem fileSystem)
	    {
	        this.dataHelper = dataHelper;
	        this.fileHelper = fileHelper;
	        this.fileSystem = fileSystem;
	    }

	    /// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public IExcelDataReader CreateBinaryReader(Stream fileStream)
		{
            IExcelDataReader reader = new ExcelBinaryReader(dataHelper);
			reader.Initialize(fileStream);

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public IExcelDataReader CreateBinaryReader(Stream fileStream, ReadOption option)
		{
            IExcelDataReader reader = new ExcelBinaryReader(dataHelper);
		    reader.ReadOption = option;
			reader.Initialize(fileStream);

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate)
		{
			IExcelDataReader reader = CreateBinaryReader(fileStream);
			((ExcelBinaryReader) reader).ConvertOaDate = convertOADate;

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate, ReadOption readOption)
		{
			IExcelDataReader reader = CreateBinaryReader(fileStream, readOption);
			((ExcelBinaryReader)reader).ConvertOaDate = convertOADate;

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelOpenXmlReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public IExcelDataReader CreateOpenXmlReader(Stream fileStream)
		{
            IExcelDataReader reader = new ExcelOpenXmlReader(fileSystem, fileHelper, dataHelper);
			reader.Initialize(fileStream);

			return reader;
		}
	}
}
