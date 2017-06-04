using System;

namespace ExcelDataReader
{
    public class DataColumn
    {
        public DataColumn(string key, Type type)
        {
            ColumnName = key;
            DataType = type;
        }

        public string ColumnName { get; set; }

        public Type DataType { get; set; }

        public string Caption { get; set; }
    }
}
