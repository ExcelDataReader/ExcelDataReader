using System;
using System.Collections.Generic;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxWorkbook
    {
        public XlsxWorkbook(List<XlsxWorksheet> sheets, XlsxSST sst, XlsxStyles styles)
        {
            Sheets = sheets;
            SST = sst;
            Styles = styles;
        }

        public List<XlsxWorksheet> Sheets { get; set; }

        public XlsxSST SST { get; }

        public XlsxStyles Styles { get; }
    }
}
