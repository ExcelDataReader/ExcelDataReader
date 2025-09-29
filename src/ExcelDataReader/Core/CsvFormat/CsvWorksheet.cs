using System.Text;

namespace ExcelDataReader.Core.CsvFormat;

internal sealed class CsvWorksheet : IWorksheet
{
    public CsvWorksheet(Stream stream, Encoding fallbackEncoding, char[] autodetectSeparators, int analyzeInitialCsvRows, char? quoteChar = null, bool trimWhiteSpace = true)
    {
        Stream = stream;
        QuoteChar = quoteChar;
        TrimWhiteSpace = trimWhiteSpace;
        Stream.Seek(0, SeekOrigin.Begin);
        try
        {
            // Try as UTF-8 first, or use BOM if present
            CsvAnalyzer.Analyze(Stream, autodetectSeparators, Encoding.UTF8, analyzeInitialCsvRows, quoteChar, trimWhiteSpace, out var fieldCount, out var separator, out var encoding, out var bomLength, out var rowCount);
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

            CsvAnalyzer.Analyze(Stream, autodetectSeparators, fallbackEncoding, analyzeInitialCsvRows, quoteChar, trimWhiteSpace, out var fieldCount, out var separator, out var encoding, out var bomLength, out var rowCount);
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

    public string CodeName => null;

    public string VisibleState => null;

    public HeaderFooter HeaderFooter => null;

    public CellRange[] MergeCells => null;

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

    public int FirstRow { get; private set; }

    public int LastRow { get; private set; }

    public int FirstColumn { get; private set; }

    public int LastColumn { get; private set; }

    public Stream Stream { get; }

    public Encoding Encoding { get; }

    public char? QuoteChar { get; }

    public char Separator { get; }

    public List<Column> ColumnWidths => null;

    public bool TrimWhiteSpace { get; }

    private int BomLength { get; set; }

    private bool AnalyzedPartial { get; }

    private int AnalyzedRowCount { get; }

    public IEnumerable<Row> ReadRows()
    {
        var bufferSize = 1024;
        var buffer = new byte[bufferSize];
        var rowIndex = 0;
        var csv = new CsvParser(Separator, Encoding, QuoteChar, TrimWhiteSpace);
        var skipBomBytes = BomLength;

        Stream.Seek(0, SeekOrigin.Begin);
        while (Stream.Position < Stream.Length)
        {
            var bytesRead = Stream.ReadAtLeast(buffer, 0, bufferSize);
            csv.ParseBuffer(buffer, skipBomBytes, bytesRead - skipBomBytes, out var bufferRows);

            skipBomBytes = 0; // Only skip bom on first iteration

            foreach (var row in GetReaderRows(rowIndex, bufferRows))
            {
                yield return row;
            }

            rowIndex += bufferRows.Count;
        }

        csv.Flush(out var flushRows);
        foreach (var row in GetReaderRows(rowIndex, flushRows))
        {
            yield return row;
        }
    }

    private static IEnumerable<Row> GetReaderRows(int rowIndex, List<List<string>> rows)
    {
        foreach (var row in rows)
        {
            var cells = new List<Cell>(row.Count);
            for (var index = 0; index < row.Count; index++)
            {
                object value = row[index];
                cells.Add(new Cell(index, value, ExtendedFormat.Zero, null));
            }

            yield return new Row(rowIndex, 12.75 /* 255 twips */, cells);

            rowIndex++;
        }
    }
}