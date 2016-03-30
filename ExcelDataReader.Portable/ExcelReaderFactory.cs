using System.IO;
using System.Threading.Tasks;
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
		public async Task<IExcelDataReader> CreateBinaryReaderAsync(Stream fileStream)
		{
            IExcelDataReader reader = new ExcelBinaryReader(dataHelper);
            await reader.InitializeAsync(fileStream);

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
        public async Task<IExcelDataReader> CreateBinaryReaderAsync(Stream fileStream, ReadOption option)
		{
            IExcelDataReader reader = new ExcelBinaryReader(dataHelper);
		    reader.ReadOption = option;
			await reader.InitializeAsync(fileStream);

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
        public async Task<IExcelDataReader> CreateBinaryReaderAsync(Stream fileStream, bool convertOADate)
		{
			IExcelDataReader reader = await CreateBinaryReaderAsync(fileStream);
			((ExcelBinaryReader) reader).ConvertOaDate = convertOADate;

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
        public async Task<IExcelDataReader> CreateBinaryReaderAsync(Stream fileStream, bool convertOADate, ReadOption readOption)
		{
			IExcelDataReader reader = await CreateBinaryReaderAsync(fileStream, readOption);
			((ExcelBinaryReader)reader).ConvertOaDate = convertOADate;

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelOpenXmlReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
        public async Task<IExcelDataReader> CreateOpenXmlReader(Stream fileStream)
		{
            IExcelDataReader reader = new ExcelOpenXmlReader(fileSystem, fileHelper, dataHelper);
			await reader.InitializeAsync(fileStream);

			return reader;
		}

        /// <summary>
        /// Creates an instance of <see cref="ExcelOpenXmlReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <returns></returns>
        public async Task<IExcelDataReader> CreateOpenXmlReader(Stream fileStream, ReadOption option)
        {
            IExcelDataReader reader = new ExcelOpenXmlReader(fileSystem, fileHelper, dataHelper);
            reader.ReadOption = option;
            await reader.InitializeAsync(fileStream);

            return reader;
        }
	}
}
