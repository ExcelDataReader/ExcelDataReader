using System.Text;
using ExcelDataReader.Core.CsvFormat;

namespace ExcelDataReader;

internal sealed class ExcelCsvReader : ExcelDataReader<CsvWorkbook, CsvWorksheet>
{
    public ExcelCsvReader(Stream stream, Encoding fallbackEncoding, char[] autodetectSeparators, int analyzeInitialCsvRows, char? quoteChar = null, bool trimWhiteSpace = true, char? escapeChar = null)
    {
        Workbook = new CsvWorkbook(stream, fallbackEncoding, autodetectSeparators, analyzeInitialCsvRows, quoteChar, trimWhiteSpace, escapeChar);

        // By default, the data reader is positioned on the first result.
        Reset();
    }

    public override void Close()
    {
        base.Close();
        Workbook?.Stream?.Dispose();
        Workbook = null;
    }
}