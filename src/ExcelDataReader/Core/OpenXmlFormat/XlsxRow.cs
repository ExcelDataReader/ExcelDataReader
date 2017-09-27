namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxRow : XlsxElement
    {
        public XlsxRow()
            : base(XlsxElementType.Row)
        {
        }

        public Row Row { get; set; }
    }
}
