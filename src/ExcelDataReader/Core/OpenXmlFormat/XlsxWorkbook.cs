using System;
using System.Collections.Generic;
using System.Xml;
using System.IO;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxWorkbook
    {

        private XlsxWorkbook() { }

        public XlsxWorkbook(List<XlsxWorksheet> sheets, XlsxSST _SST, XlsxStyles _Styles)
        {
            this.sheets = sheets;
            this._SST = _SST;
            this._Styles = _Styles;
        }

        private List<XlsxWorksheet> sheets;

        public List<XlsxWorksheet> Sheets
        {
            get { return sheets; }
            set { sheets = value; }
        }

        private XlsxSST _SST;

        public XlsxSST SST
        {
            get { return _SST; }
        }

        private XlsxStyles _Styles;

        public XlsxStyles Styles
        {
            get { return _Styles; }
        }

    }
}
