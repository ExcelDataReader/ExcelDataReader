using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Excel
{
	public static class Factory
	{
		public static IExcelDataReader CreateReader(Stream fileStream, ExcelFileType excelFileType)
		{
			IExcelDataReader reader = null;

			switch (excelFileType)
			{
				case ExcelFileType.Binary:
					reader = new ExcelBinaryReader();
					reader.Initialize(fileStream);
					break;
				case ExcelFileType.OpenXml:
					reader = new ExcelOpenXmlReader();
					reader.Initialize(fileStream);
					break;
				default:
					break;
			}

			return reader;
		}

	}
}
