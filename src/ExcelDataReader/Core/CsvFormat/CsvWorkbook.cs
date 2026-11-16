using System.Text;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core.CsvFormat;

internal sealed class CsvWorkbook(Stream stream, Func<Encoding> fallbackEncoding, string[] autodetectSeparators, int analyzeInitialCsvRows, char? quoteChar = null, bool trimWhiteSpace = true, char? escapeChar = null) : IWorkbook<CsvWorksheet>
{
    public int ResultsCount => 1;

    public int ActiveSheet => 0;

    public Stream Stream { get; } = stream;

    public char? QuoteChar { get; } = quoteChar;

    public string[] AutodetectSeparators { get; } = autodetectSeparators;

    public int AnalyzeInitialCsvRows { get; } = analyzeInitialCsvRows;

    public bool TrimWhiteSpace { get; } = trimWhiteSpace;

    public char? EscapeChar { get; } = escapeChar;

    public IEnumerable<CsvWorksheet> ReadWorksheets()
    {
        yield return new CsvWorksheet(Stream, fallbackEncoding, AutodetectSeparators, AnalyzeInitialCsvRows, QuoteChar, TrimWhiteSpace, EscapeChar);
    }

    public NumberFormatString? GetNumberFormatString(int index, IFormatProvider? provider) => null;

    public void Dispose()
    {
        Stream.Dispose();
    }
}