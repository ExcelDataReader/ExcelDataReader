using System;
using ExcelDataReader.Portable.Data;

namespace Excel
{
	public class DataHelper : IDataHelper
	{
		public bool IsDBNull(object value)
		{
			return value != null && value.Equals(DBNull.Value);
		}
	}
}
