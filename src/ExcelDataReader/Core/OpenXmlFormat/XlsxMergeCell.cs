namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxMergeCell : XlsxElement
    {
        public XlsxMergeCell()
            : base(XlsxElementType.MergeCell)
        {
        }

        public CellRange Value { get; set; }
    }
}
