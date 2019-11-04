using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class CellRecord : Record
    {
        public CellRecord(int columnIndex, int xfIndex, object value)
        {
            ColumnIndex = columnIndex;
            XfIndex = xfIndex;
            Value = value;
        }

        public int ColumnIndex { get; }

        public int XfIndex { get; }

        public object Value { get; }
    }
}
