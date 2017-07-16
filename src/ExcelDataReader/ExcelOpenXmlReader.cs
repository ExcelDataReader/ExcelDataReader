using System;
using System.IO;
using ExcelDataReader.Core;
using ExcelDataReader.Core.OpenXmlFormat;

namespace ExcelDataReader
{
    internal class ExcelOpenXmlReader : ExcelDataReader<XlsxWorkbook, XlsxWorksheet>
    {
        public ExcelOpenXmlReader(Stream stream, ExcelReaderConfiguration configuration)
            : base(configuration)
        {
            Document = new ZipWorker(stream);
            Workbook = new XlsxWorkbook(Document);

            // By default, the data reader is positioned on the first result.
            Reset();
        }

        private ZipWorker Document { get; set; }

        public override void Close()
        {
            base.Close();

            Document?.Dispose();
            Workbook = null;
            Document = null;
        }
    }
}
