using System.Collections;
using System.Collections.Generic;

namespace ExcelDataReader
{
    public class DataTableCollection : IEnumerable<DataTable>
    {
        private readonly List<DataTable> tables = new List<DataTable>();

        public int Count => tables.Count;

        public DataTable this[int index] => tables[index];

        public DataTable this[string name]
        {
            get
            {
                foreach (var table in tables)
                {
                    if (table.TableName == name)
                        return table;
                }

                return null;
            }
        }

        public void Add(DataTable table)
        {
            tables.Add(table);
        }

        public void AddRange(DataTable[] range)
        {
            tables.AddRange(range);
        }

        public void Clear()
        {
            tables.Clear();
        }

        public IEnumerator<DataTable> GetEnumerator()
        {
            return tables.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return tables.GetEnumerator();
        }
    }
}
