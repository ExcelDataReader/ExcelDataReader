using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class WorkbookActRecord : Record
    {
        public WorkbookActRecord(int activeSheet)
        {
            this.ActiveSheet = activeSheet;
        }

        public int ActiveSheet { get; }
    }
}
