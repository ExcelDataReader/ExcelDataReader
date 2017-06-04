#if NETSTANDARD1_3

using System;
using System.Collections.Generic;

namespace ExcelDataReader
{
    public class DataSet
    {
        public DataSet()
        {
            Tables = new DataTableCollection();
        }

        public DataTableCollection Tables { get; set; }

        public void AcceptChanges()
        {
        }
    }
}

#endif
