using System.IO;
using System.Text;
using ExcelDataReader.Core.CsvFormat;

namespace ExcelDataReader
{
    internal class ExcelCsvReader : ExcelDataReader<CsvWorkbook, CsvWorksheet>
    {
        public ExcelCsvReader(Stream stream, Encoding fallbackEncoding, char[] autodetectSeparators)
        {
            Workbook = new CsvWorkbook(stream, fallbackEncoding, autodetectSeparators);

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
