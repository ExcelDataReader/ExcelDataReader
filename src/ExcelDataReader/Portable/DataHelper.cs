using System;
using ExcelDataReader.Data;

namespace Excel
{
	public class DataHelper : IDataHelper {
		public bool IsDBNull(object value) {
#if NET20 || NET45
			return Convert.IsDBNull(value);
#else
			return value != null && value.Equals(DBNull.Value);
#endif
		}
	}
}
