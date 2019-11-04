using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class RowHeaderRecord : Record
    {
        public RowHeaderRecord(int rowIndex, bool hidden, double? height) 
        {
            RowIndex = rowIndex;
            Hidden = hidden;
            Height = height;
        }

        public int RowIndex { get; }

        public bool Hidden { get; }

        public double? Height { get; }
    }
}
