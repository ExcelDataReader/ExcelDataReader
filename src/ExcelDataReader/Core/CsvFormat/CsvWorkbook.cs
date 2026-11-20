using System.Text;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core.CsvFormat;

internal sealed class CsvWorkbook(Stream stream, Func<Encoding> fallbackEncoding, string[] autodetectSeparators, int analyzeInitialCsvRows, char? quoteChar = null, bool trimWhiteSpace = true, char? escapeChar = null) : IWorkbook<CsvWorksheet>
{
    public int ResultsCount => 1;

    public int ActiveSheet => 0;

    private Stream Stream { get; } = stream;

    private char? QuoteChar { get; } = quoteChar;

    private string[] AutodetectSeparators { get; } = autodetectSeparators;

    private int AnalyzeInitialCsvRows { get; } = analyzeInitialCsvRows;

    private bool TrimWhiteSpace { get; } = trimWhiteSpace;

    private char? EscapeChar { get; } = escapeChar;

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