using System;
using System.IO;
using ExcelDataReader.Data;

namespace Excel
{
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
			var reader = new ExcelBinaryReader();
			reader.Initialize(fileStream);
			return reader;
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
            var reader = new ExcelBinaryReader();
			reader.ReadOption = option;
			reader.Initialize(fileStream);
			return reader;
		}

	    /// <summary>
	    /// Creates an instance of <see cref="ExcelBinaryReader"/>
	    /// </summary>
	    /// <param name="fileStream">The file stream.</param>
	    /// <param name="convertOADate"></param>
	    /// <returns></returns>
	    public static IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate)
		{
			var reader = new ExcelBinaryReader();
			reader.ConvertOaDate = convertOADate;
			reader.Initialize(fileStream);
			return reader;
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
			var reader = new ExcelBinaryReader();
			reader.ConvertOaDate = convertOADate;
			reader.ReadOption = readOption;
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
			var reader = new ExcelOpenXmlReader();
			reader.Initialize(fileStream);
			return reader;
		}
	}
}
