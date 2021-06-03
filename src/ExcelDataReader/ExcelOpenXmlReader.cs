using System.Collections.Generic;
using System.IO;
using ExcelDataReader.Core;
using ExcelDataReader.Core.OpenXmlFormat;

namespace ExcelDataReader
{
    internal class ExcelOpenXmlReader : ExcelDataReader<XlsxWorkbook, XlsxWorksheet>
    {
        public ExcelOpenXmlReader(Stream stream, Dictionary<string, object> runtimeInjectedFields = null)

        {
            Document = new ZipWorker(stream);
            Workbook = new XlsxWorkbook(Document);
            InjectedFields = runtimeInjectedFields; // useful in a bulk insert stream; we may optionally wish to inject fixed column identifier values in the stream

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
