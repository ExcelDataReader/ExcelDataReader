using System;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal sealed class XlsxHeaderFooter : XlsxElement
    {
        public XlsxHeaderFooter(HeaderFooter headerFooter)
            : base(XlsxElementType.HeaderFooter)
        {
            Value = headerFooter;
        }
        
        public HeaderFooter Value { get; }
    }
}