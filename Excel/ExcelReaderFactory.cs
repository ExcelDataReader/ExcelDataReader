using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

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
			IExcelDataReader reader = new ExcelBinaryReader();
			reader.Initialize(fileStream);

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public static IExcelDataReader CreateBinaryReader(Stream fileStream, ReadOption option)
		{
			IExcelDataReader reader = new ExcelBinaryReader(option);
			reader.Initialize(fileStream);

			return reader;
		}

		/// <summary>
		/// Creates an instance of <see cref="ExcelBinaryReader"/>
		/// </summary>
		/// <param name="fileStream">The file stream.</param>
		/// <returns></returns>
		public static IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate)
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
		public static IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate, ReadOption readOption)
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
		public static IExcelDataReader CreateOpenXmlReader(Stream fileStream)
		{
			IExcelDataReader reader = new ExcelOpenXmlReader();
			reader.Initialize(fileStream);

			return reader;
		}
	}
}
