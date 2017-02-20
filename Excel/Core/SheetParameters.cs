using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Core
{
    public class SheetParameters
    {
        public string SheetName { get; set; }
        public int SkipRows { get; set; }
        public bool IsFirstRowAsColumnNames { get; set; }

        private SheetParameters() { }
        public SheetParameters(string sheetName, bool isFirstRowAsColumnNames = true, int skipRows = 0)
        {
            SheetName = sheetName;
            IsFirstRowAsColumnNames = isFirstRowAsColumnNames;
            SkipRows = skipRows;
        }

    }
}
