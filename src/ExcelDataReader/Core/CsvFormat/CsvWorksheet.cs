using System.Text;

namespace ExcelDataReader.Core.CsvFormat;

internal sealed class CsvWorksheet : IWorksheet
{
    public CsvWorksheet(Stream stream, Encoding fallbackEncoding, string[] autodetectSeparators, int analyzeInitialCsvRows, char? quoteChar = null, bool trimWhiteSpace = true, char? escapeChar = null)
    {
        Stream = stream;
        QuoteChar = quoteChar;
        TrimWhiteSpace = trimWhiteSpace;
        EscapeChar = escapeChar;
        Stream.Seek(0, SeekOrigin.Begin);
        try
        {
            // Try as UTF-8 first, or use BOM if present
            CsvAnalyzer.Analyze(Stream, autodetectSeparators, Encoding.UTF8, analyzeInitialCsvRows, quoteChar, trimWhiteSpace, escapeChar, out var fieldCount, out var separator, out var encoding, out var bomLength, out var rowCount);
            FieldCount = fieldCount;
            AnalyzedRowCount = rowCount;
            AnalyzedPartial = analyzeInitialCsvRows > 0;
            Encoding = encoding;
            Separator = separator;
            BomLength = bomLength;
        }
        catch (DecoderFallbackException)
        {
            // If cannot parse as UTF-8, try fallback encoding
            Stream.Seek(0, SeekOrigin.Begin);

            CsvAnalyzer.Analyze(Stream, autodetectSeparators, fallbackEncoding, analyzeInitialCsvRows, quoteChar, trimWhiteSpace, escapeChar, out var fieldCount, out var separator, out var encoding, out var bomLength, out var rowCount);
            FieldCount = fieldCount;
            AnalyzedRowCount = rowCount;
            AnalyzedPartial = analyzeInitialCsvRows > 0;
            Encoding = encoding;
            Separator = separator;
            BomLength = bomLength;
        }
    }

    public static int ActiveSheet => 0;

    public static bool IsActiveSheet => true;

    public string Name => string.Empty;

    public string? CodeName => null;

    public string? VisibleState => null;

    public HeaderFooter? HeaderFooter => null;

    public CellRange[] MergeCells => [];

    public int FieldCount { get; }

    public int RowCount
    {
        get
        {
            if (AnalyzedPartial)
            {
                throw new InvalidOperationException("Cannot use RowCount with AnalyzeInitialCsvRows > 0");
            }

            return AnalyzedRowCount;
        }
    }

    public CellRange? Dimension => null;

    public Stream Stream { get; }

    public Encoding Encoding { get; }

    public char? QuoteChar { get; }

    public string Separator { get; }

    public char? EscapeChar { get; }

    public List<Column> ColumnWidths => [];

    public bool TrimWhiteSpace { get; }

    public double DefaultRowHeight => 12.75;

    private int BomLength { get; set; }

    private bool AnalyzedPartial { get; }

    private int AnalyzedRowCount { get; }

    public Dictionary<long, CellCommentThread>? LoadComments() => null;

    public IWorksheetReader OpenReader() => new CsvWorksheetReader(this, BomLength);
}