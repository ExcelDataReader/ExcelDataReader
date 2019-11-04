using System;
using System.IO;

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class WorkbookPrRecord : Record
    {
        public WorkbookPrRecord(bool date1904)
        {
            Date1904 = date1904;
        }

        public bool Date1904 { get; }
    }
}