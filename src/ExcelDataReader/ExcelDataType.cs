using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader
{
    /// <summary>
    /// Represent a column data type.
    /// </summary>
    public class ExcelDataType
    {
        public string ColumnName { get; set; }
        public Type ColumnType { get; set; }
    }
}
