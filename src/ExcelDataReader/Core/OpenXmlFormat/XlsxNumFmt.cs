namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxNumFmt
    {
        public XlsxNumFmt(int id, string formatCode)
        {
            Id = id;
            FormatCode = formatCode;
        }

        public int Id { get; set; }

        public string FormatCode { get; set; }
    }
}
