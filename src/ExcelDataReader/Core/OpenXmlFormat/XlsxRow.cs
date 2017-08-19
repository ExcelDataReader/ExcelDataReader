using System;
using System.Collections.Generic;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxRow : XlsxElement
    {
        public XlsxRow()
            : base(XlsxElementType.Row)
        {
        }

        public int RowIndex { get; set; }

        public double RowHeight { get; set; }

        public List<XlsxCell> Cells { get; set; } = new List<XlsxCell>();

        public int GetMaxColumnIndex()
        {
            int columnIndex = int.MinValue;
            foreach (var cell in Cells)
            {
                columnIndex = Math.Max(cell.ColumnIndex, columnIndex);
            }

            return columnIndex;
        }
    }
}
