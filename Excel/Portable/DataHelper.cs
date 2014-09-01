using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader.Data;

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
