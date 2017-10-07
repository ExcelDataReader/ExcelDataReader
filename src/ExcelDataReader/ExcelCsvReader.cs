using System.IO;
using ExcelDataReader.Core.CsvFormat;

namespace ExcelDataReader
{
    internal class ExcelCsvReader : ExcelDataReader<CsvWorkbook, CsvWorksheet>
    {
        public ExcelCsvReader(Stream stream, ExcelReaderConfiguration configuration)
            : base(configuration)
        {
            Workbook = new CsvWorkbook(stream, Configuration.FallbackEncoding);

            // By default, the data reader is positioned on the first result.
            Reset();
        }

        public override void Close()
        {
            base.Close();
            Workbook.Stream?.Dispose();
            Workbook = null;
        }
    }
}
