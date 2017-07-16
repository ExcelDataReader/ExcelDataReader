// ReSharper disable InconsistentNaming
namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxXf
    {
        public XlsxXf(int id, int numFmtId, string applyNumberFormat)
        {
            Id = id;
            NumFmtId = numFmtId;
            ApplyNumberFormat = applyNumberFormat != null && applyNumberFormat == "1";
        }

        public int Id { get; set; }

        public int NumFmtId { get; set; }

        public bool ApplyNumberFormat { get; set; }
    }
}
