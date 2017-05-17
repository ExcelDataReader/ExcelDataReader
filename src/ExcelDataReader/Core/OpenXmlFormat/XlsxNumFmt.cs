namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxNumFmt
    {
        public const string NNumFmt = "numFmt";
        public const string ANumFmtId = "numFmtId";
        public const string AFormatCode = "formatCode";

        public XlsxNumFmt(int id, string formatCode)
        {
            Id = id;
            FormatCode = formatCode;
        }

        public int Id { get; set; }

        public string FormatCode { get; set; }
    }
}
