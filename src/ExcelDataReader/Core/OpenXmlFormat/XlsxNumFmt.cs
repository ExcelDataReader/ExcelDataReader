using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxNumFmt
    {
        public XlsxNumFmt(int id, string formatCode)
        {
            Id = id;
            FormatCode = new NumberFormatString(formatCode);
        }

        public int Id { get; set; }

        public NumberFormatString FormatCode { get; set; }
    }
}
