using System;
using System.Collections;
using System.Collections.Generic;

namespace ExcelDataReader
{
    public class DataColumnCollection : IEnumerable<DataColumn>
    {
        private readonly List<DataColumn> columns = new List<DataColumn>();

        public int Count => columns.Count;

        public DataColumn this[int index] => columns[index];

        public DataColumn this[string name]
        {
            get
            {
                foreach (var column in columns)
                {
                    if (column.ColumnName == name)
                        return column;
                }

                return null;
            }
        }

        public void Add(string key, Type type)
        {
            Add(new DataColumn(key, type));
        }

        public void Add(DataColumn column)
        {
            if (column.ColumnName != null && this[column.ColumnName] != null)
                throw new ArgumentException("Duplicate column name: " + column.ColumnName);

            columns.Add(column);
        }

        public IEnumerator<DataColumn> GetEnumerator()
        {
            return columns.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return columns.GetEnumerator();
        }
    }
}
