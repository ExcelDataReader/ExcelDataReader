using System.Collections;
using System.Collections.Generic;

namespace ExcelDataReader
{
    public class DataRowCollection : IEnumerable<DataRow>
    {
        private readonly List<DataRow> rows = new List<DataRow>();

        public int Count => rows.Count;

        public DataRow this[int index] => rows[index];

        public void Add(params object[] values)
        {
            rows.Add(new DataRow(values));
        }

        public void Add(DataRow row)
        {
            rows.Add(row);
        }

        public void Clear()
        {
            rows.Clear();
        }

        public IEnumerator<DataRow> GetEnumerator()
        {
            return rows.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return rows.GetEnumerator();
        }
    }
}
