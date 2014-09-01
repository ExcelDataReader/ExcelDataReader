using System;
using ExcelDataReader.Portable.Data;

namespace Excel.Portable
{
    internal class DataHelper : IDataHelper
    {
        public bool IsDBNull(object value)
        {
            return Convert.IsDBNull(value);
        }
    }
}
