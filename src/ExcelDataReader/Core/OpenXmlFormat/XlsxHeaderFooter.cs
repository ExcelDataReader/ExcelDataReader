using System;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal sealed class XlsxHeaderFooter : XlsxElement
    {
        public XlsxHeaderFooter(bool isHeader, string value)
            : base((XlsxElementType)XlsxElementType.HeaderFooter)
        {
            IsHeader = isHeader;
            Value = value;
        }
        
        public bool IsHeader { get; }

        public string Value { get; }
    }
}