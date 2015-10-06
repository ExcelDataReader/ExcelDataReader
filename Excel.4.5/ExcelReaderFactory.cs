using System;
using System.IO;
using ExcelDataReader.Desktop.Portable;
using ExcelDataReader.Portable.Async;
using ExcelDataReader.Portable.Data;
using ExcelDataReader.Portable.IO;
using ExcelDataReader.Portable.Misc;
using PCLStorage;

namespace Excel
{
	/// <summary>
	/// The ExcelReader Factory
	/// </summary>
	public static class ExcelReaderFactory
	{
	    private static readonly IDataHelper dataHelper = new DataHelper();
	    private static readonly IFileHelper fileHelper = new FileHelper();
	    private static readonly IFileSystem fileSystem = FileSystem.Current;
		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public static IExcelDataReader CreateBinaryReader(Stream fileStream)
		{
            var factory = CreateFactory();

		    var reader = AsyncHelper.RunSync(() => factory.CreateBinaryReaderAsync(fileStream));

            return new ExcelBinaryReader(reader);
		}

	    private static ExcelDataReader.Portable.ExcelReaderFactory CreateFactory()
	    {
	        return new ExcelDataReader.Portable.ExcelReaderFactory(dataHelper, fileHelper, fileSystem);
	    }

        /// <summary>
        /// Creates an instance of <see cref="ExcelBinaryReader"/> or <see cref="ExcelOpenXmlReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <returns></returns>
        public static IExcelDataReader CreateReader(Stream fileStream)
        {
            const ulong xlsSignature = 0xE11AB1A1E011CFD0;
            var buf = new byte[512];
            fileStream.Seek(0, SeekOrigin.Begin);
            fileStream.Read(buf, 0, 512);
            fileStream.Seek(0, SeekOrigin.Begin);

            var hdr = BitConverter.ToUInt64(buf, 0x0);

            if (hdr == xlsSignature)
                return CreateBinaryReader(fileStream);
            return CreateOpenXmlReader(fileStream);

        }



        /// <summary>
        /// Creates an instance of <see cref="ExcelBinaryReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <returns></returns>
        public static IExcelDataReader CreateBinaryReader(Stream fileStream, ReadOption option)
		{
            var factory = CreateFactory();

            var portableReadOption = (ExcelDataReader.Portable.ReadOption)option;
            var reader = AsyncHelper.RunSync(() => factory.CreateBinaryReaderAsync(fileStream, portableReadOption));

            return new ExcelBinaryReader(reader);
		}

	    /// <summary>
	    /// Creates an instance of <see cref="ExcelBinaryReader"/>
	    /// </summary>
	    /// <param name="fileStream">The file stream.</param>
	    /// <param name="convertOADate"></param>
	    /// <returns></returns>
	    public static IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate)
		{
            var factory = CreateFactory();

            var reader = AsyncHelper.RunSync(() => factory.CreateBinaryReaderAsync(fileStream, convertOADate));

            return new ExcelBinaryReader(reader);

		}

	    /// <summary>
	    /// Creates an instance of <see cref="ExcelBinaryReader"/>
	    /// </summary>
	    /// <param name="fileStream">The file stream.</param>
	    /// <param name="convertOADate"></param>
	    /// <param name="readOption"></param>
	    /// <returns></returns>
	    public static IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate, ReadOption readOption)
		{
            var factory = CreateFactory();

            var portableReadOption = (ExcelDataReader.Portable.ReadOption)readOption;
            var reader = AsyncHelper.RunSync(() => factory.CreateBinaryReaderAsync(fileStream, convertOADate, portableReadOption));

            return new ExcelBinaryReader(reader);
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelOpenXmlReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public static IExcelDataReader CreateOpenXmlReader(Stream fileStream)
		{
            var factory = CreateFactory();

			var reader = AsyncHelper.RunSync(() => factory.CreateOpenXmlReader(fileStream));

			return new ExcelOpenXmlReader(reader);
		}
	}
}
