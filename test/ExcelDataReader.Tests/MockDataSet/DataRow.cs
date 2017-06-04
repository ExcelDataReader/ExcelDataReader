using System;

namespace ExcelDataReader
{
    public class DataRow
    {
        readonly object[] values;

        internal DataRow(object[] itemArray)
        {
            values = (object[])itemArray.Clone();
        }

        public object[] ItemArray { 
            get {
                var result = new object[values.Length];
                for (var i = 0; i < values.Length; i++) {
                    result[i] = values[i] ?? DBNull.Value;
                }
                return result;
            }
        }

        public object this[int index]
        {
            get
            {
                var result = values[index];
                return result ?? DBNull.Value;
            }

            set => values[index] = value;
        }

        public bool IsNull(int i)
        {
            return ItemArray[i] == DBNull.Value;
        }
    }
}
