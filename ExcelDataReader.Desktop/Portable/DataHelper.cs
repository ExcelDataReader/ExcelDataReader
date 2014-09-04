using System;
using ExcelDataReader.Portable.Data;

namespace ExcelDataReader.Desktop.Portable
{
    public class DataHelper : IDataHelper
    {
        public bool IsDBNull(object value)
        {
            return Convert.IsDBNull(value);
        }
    }
}
