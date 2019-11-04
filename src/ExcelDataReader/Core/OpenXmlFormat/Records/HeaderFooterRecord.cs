using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class HeaderFooterRecord : Record
    {
        public HeaderFooterRecord(HeaderFooter headerFooter) 
        {
            HeaderFooter = headerFooter;
        }

        public HeaderFooter HeaderFooter { get; }
    }
}
