#if NETSTANDARD1_3
using System;
using System.Linq;

namespace Excel
{
    public class DataRow
    {
        internal DataRow(object[] itemArray)
        {
            ItemArray = itemArray;
        }

        public object[] ItemArray { get; }

        public object this[int index]
        {
            get
            {
                var result = ItemArray[index];
                return result ?? DBNull.Value;
            }

            set => ItemArray[index] = value;
        }

        public bool IsNull(int i)
        {
            return ItemArray[i] == null;
        }
    }
}
#endif