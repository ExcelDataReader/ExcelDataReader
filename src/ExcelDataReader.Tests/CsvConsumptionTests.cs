using System.Text;
using ExcelDataReader.Core;
using ExcelDataReader.Core.CsvFormat;

namespace ExcelDataReader.Tests;

public class CsvConsumptionTests
{
    [TestCase(',', '"', '\\', true)]
    [TestCase('\t', '"', '"', true)]
    [TestCase('|', null, null, false)]
    [TestCase(',', '\r', '\n', true)]
    public void RunScanningMatchesCharacterSizedBuffers(char separator, char? quote, char? escape, bool trim)
    {
        var random = new Random(4721);
        const string alphabet = "ab \t,\"|\\\r\né終";
        for (int sample = 0; sample < 40; sample++)
        {
            var text = new char[random.Next(50, 2500)];
            for (int i = 0; i < text.Length; i++)
                text[i] = alphabet[random.Next(alphabet.Length)];
            byte[] bytes = Encoding.UTF8.GetBytes(text);
            foreach (var mode in new[] { WorksheetReadMode.Materialize, WorksheetReadMode.Scan })
            {
                var expected = Parse(1, mode);
                Assert.That(Parse(1024, mode), Is.EqualTo(expected), $"Sample {sample}, mode {mode}");
            }

            List<string> Parse(int bufferSize, WorksheetReadMode mode)
            {
                var parser = new CsvParser(separator, Encoding.UTF8, quote, trim, escape, mode);
                var handler = new ParserHandler(pause: true);
                for (int start = 0; start < bytes.Length; start += bufferSize)
                {
                    int offset = start;
                    int count = Math.Min(bufferSize, bytes.Length - start);
                    while (parser.ParseBuffer(bytes, ref offset, ref count, ref handler) == WorksheetReadStatus.Paused)
                    {
                    }
                }

                while (parser.Flush(ref handler) == WorksheetReadStatus.Paused)
                {
                }

                return handler.Events;
            }
        }
    }

    [TestCase(1)]
    [TestCase(7)]
    [TestCase(1024)]
    public void RunsPreserveTrailingWhitespaceAcrossBuffersAndQuotedTransitions(int bufferSize)
    {
        string padding = new(' ', 1100);
        string text = "\"a  \"" + padding + ",\"b\" \tX \t,\"c\t \",d\t \r\n";
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        var parser = new CsvParser(',', Encoding.UTF8, '"');
        var handler = new ParserHandler(pause: true);
        for (int start = 0; start < bytes.Length; start += bufferSize)
        {
            int offset = start;
            int count = Math.Min(bufferSize, bytes.Length - start);
            while (parser.ParseBuffer(bytes, ref offset, ref count, ref handler) == WorksheetReadStatus.Paused)
            {
            }
        }

        while (parser.Flush(ref handler) == WorksheetReadStatus.Paused)
        {
        }

        Assert.That(handler.Events, Is.EqualTo(new[]
        {
            "row:0", "field:0:a  ", "field:1:b \tX", "field:2:c\t ", "field:3:d", "end:4:False:False",
        }));
    }

    [TestCase(1)]
    [TestCase(2)]
    [TestCase(7)]
    [TestCase(1024)]
    public void ParserResumesEveryCallbackAcrossByteAndCharacterBoundaries(int bufferSize)
    {
        string longValue = new('x', 1021);
        string text = "  a  ,\"" + longValue + "é\"\"z\\\"q\r\nend\", \t\r\n\rnext,last,";
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        var parser = new CsvParser(',', Encoding.UTF8, '"', escapeChar: '\\');
        var handler = new ParserHandler(pause: true);

        for (int start = 0; start < bytes.Length; start += bufferSize)
        {
            int offset = start;
            int count = Math.Min(bufferSize, bytes.Length - start);
            int callbacks = 0;
            while (parser.ParseBuffer(bytes, ref offset, ref count, ref handler) == WorksheetReadStatus.Paused)
                Assert.That(++callbacks, Is.LessThan(100));
            Assert.That(count, Is.Zero);
        }

        int flushCallbacks = 0;
        while (parser.Flush(ref handler) == WorksheetReadStatus.Paused)
            Assert.That(++flushCallbacks, Is.LessThan(10));

        Assert.That(handler.Events, Is.EqualTo(new[]
        {
            "row:0", "field:0:a", "field:1:" + longValue + "é\"z\"q\r\nend", "field:2:", "end:3:False:False",
            "row:1", "field:0:", "end:1:True:False",
            "row:2", "field:0:next", "field:1:last", "field:2:", "end:3:False:True",
        }));
    }

    [TestCase("")]
    [TestCase("\"\"")]
    [TestCase("a,b\r\n\n\"quoted \"\"text\"\",c\nlast,")]
    [TestCase("\"\" ")]
    public void ScanUsesTheSameGrammarWithoutMaterializingFields(string text)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        List<string>[] rowEvents = new List<string>[2];
        for (int i = 0; i < rowEvents.Length; i++)
        {
            var mode = i == 0 ? WorksheetReadMode.Materialize : WorksheetReadMode.Scan;
            var parser = new CsvParser(',', Encoding.UTF8, '"', mode: mode);
            var handler = new ParserHandler(pause: false);
            int offset = 0;
            int count = bytes.Length;
            parser.ParseBuffer(bytes, ref offset, ref count, ref handler);
            parser.Flush(ref handler);
            if (mode == WorksheetReadMode.Scan)
                Assert.That(handler.Events.Any(e => e.StartsWith("field:", StringComparison.Ordinal)), Is.False);
            rowEvents[i] = handler.Events.Where(e => !e.StartsWith("field:", StringComparison.Ordinal)).ToList();
        }

        Assert.That(rowEvents[1], Is.EqualTo(rowEvents[0]));
    }

    [TestCase("", 0, 0)]
    [TestCase("\"\"", 0, 0)]
    [TestCase("\"\" ", 1, 1)]
    [TestCase("\"\" \n", 0, 0)]
    [TestCase("\n,\r\n \n\"\" \nlast,", 2, 2)]
    [TestCase("x\r", 1, 1)]
    [TestCase("\r", 0, 0)]
    [TestCase("\r\n", 0, 0)]
    public void AnalysisPreservesBufferedAndFlushedEmptyRowRules(string text, int fields, int rows)
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(text));
        CsvAnalyzer.Analyze(stream, [','], Encoding.UTF8, 0, '"', true, null, out int fieldCount, out _, out _, out _, out int rowCount);
        Assert.That(fieldCount, Is.EqualTo(fields));
        Assert.That(rowCount, Is.EqualTo(rows));
    }

    [Test]
    public void AnalysisPreservesSamplingAndSeparatorTieBreaking()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("a,b\nc,d\ne,f\ng,h\nlater,has,many,fields\n"));
        CsvAnalyzer.Analyze(stream, [',', ';'], Encoding.UTF8, 1, '"', true, null, out int fields, out char separator, out _, out _, out int rows);
        Assert.That(fields, Is.EqualTo(2));
        Assert.That(separator, Is.EqualTo(','));
        Assert.That(rows, Is.EqualTo(-1));

        stream.Position = 0;
        CsvAnalyzer.Analyze(stream, [',', ';'], Encoding.UTF8, 0, '"', true, null, out fields, out separator, out _, out _, out rows);
        Assert.That(fields, Is.EqualTo(4));
        Assert.That(rows, Is.EqualTo(5));

        using var tied = new MemoryStream(Encoding.UTF8.GetBytes("a,b;c;d\n"));
        CsvAnalyzer.Analyze(tied, [',', ';'], Encoding.UTF8, 0, '"', true, null, out fields, out separator, out _, out _, out rows);
        Assert.That(separator, Is.EqualTo(';'));
        Assert.That(fields, Is.EqualTo(3));
    }

    [TestCase("utf8")]
    [TestCase("utf16")]
    [TestCase("utf16be")]
    public void WorksheetPausesBeforeCellsAndKeepsDecodedBuffersAndBom(string encodingName)
    {
        Encoding encoding = encodingName switch
        {
            "utf16" => Encoding.Unicode,
            "utf16be" => Encoding.BigEndianUnicode,
            _ => new UTF8Encoding(true),
        };
        string value = new string('x', 1022) + "é終";
        byte[] bytes = encoding.GetPreamble().Concat(encoding.GetBytes(value + ",b\r\nlast,")).ToArray();
        using var stream = new MemoryStream(bytes);
        var worksheet = new CsvWorksheet(stream, Encoding.UTF8, [','], 0, '"');
        using (var reader = worksheet.OpenReader())
        {
            var handler = new WorksheetHandler();
            Assert.That(reader.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.Paused));
            Assert.That(handler.Events, Is.EqualTo(new[] { "begin" }));
            Assert.That(reader.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.Paused));
            Assert.That(handler.Events, Is.EqualTo(new[] { "begin", "row:0" }));
            int pauses = 0;
            while (reader.Consume(ref handler) == WorksheetReadStatus.Paused)
                Assert.That(++pauses, Is.LessThan(20));
            Assert.That(handler.Events, Is.EqualTo(new[]
            {
                "begin", "row:0", "cell:0:" + value, "cell:1:b", "end",
                "row:1", "cell:0:last", "cell:1:", "end",
            }));
            Assert.That(reader.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.EndOfInput));
        }

        Assert.That(stream.CanRead, Is.True);
        using var reopened = worksheet.OpenReader();
        var restarted = new WorksheetHandler();
        reopened.Consume(ref restarted);
        reopened.Consume(ref restarted);
        Assert.That(restarted.Events, Is.EqualTo(new[] { "begin", "row:0" }));
    }

    [Test]
    public void WorksheetRetriesInvalidUtf8WithFallbackEncoding()
    {
        using var stream = new MemoryStream([0xe9, (byte)',', (byte)'b', (byte)'\n']);
        var worksheet = new CsvWorksheet(stream, Encoding.GetEncoding(1252), [','], 0);
        using var reader = worksheet.OpenReader();
        var handler = new WorksheetHandler();
        while (reader.Consume(ref handler) == WorksheetReadStatus.Paused)
        {
        }

        Assert.That(handler.Events, Is.EqualTo(new[] { "begin", "row:0", "cell:0:é", "cell:1:b", "end" }));
    }

    [Test]
    public void RowEndPauseReturnsOneCompleteRowWithoutStartingTheNext()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("a,b\nc,d\n"));
        var worksheet = new CsvWorksheet(stream, Encoding.UTF8, [','], 0);
        using var reader = worksheet.OpenReader();
        var handler = new WorksheetHandler(pauseEveryEvent: false);
        Assert.That(reader.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.Paused));
        Assert.That(handler.Events, Is.EqualTo(new[] { "begin", "row:0", "cell:0:a", "cell:1:b", "end" }));
        Assert.That(reader.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.Paused));
        Assert.That(handler.Events, Is.EqualTo(new[]
        {
            "begin", "row:0", "cell:0:a", "cell:1:b", "end",
            "row:1", "cell:0:c", "cell:1:d", "end",
        }));
        Assert.That(reader.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.EndOfInput));
    }

    private readonly struct ParserHandler(bool pause) : ICsvHandler
    {
        public List<string> Events { get; } = [];

        public WorksheetReadControl OnRow(int index) => Record("row:" + index);

        public WorksheetReadControl OnField(int index, string value) => Record("field:" + index + ":" + value);

        public WorksheetReadControl OnRowEnd(int fieldCount, bool empty, bool flushing) => Record($"end:{fieldCount}:{empty}:{flushing}");

        private WorksheetReadControl Record(string value)
        {
            Events.Add(value);
            return pause ? WorksheetReadControl.Pause : WorksheetReadControl.Continue;
        }
    }

    private struct WorksheetHandler(bool pauseEveryEvent) : IWorksheetHandler
    {
        public WorksheetHandler()
            : this(true)
        {
        }

        public List<string> Events { get; } = [];

        public WorksheetReadControl OnDimension(CellRange range) => throw new InvalidOperationException();

        public WorksheetReadControl OnCodeName(string codeName) => throw new InvalidOperationException();

        public WorksheetReadControl OnDefaultRowHeight(double? height) => throw new InvalidOperationException();

        public WorksheetReadControl OnColumn(Column column) => throw new InvalidOperationException();

        public WorksheetReadControl OnMergeCell(CellRange range) => throw new InvalidOperationException();

        public WorksheetReadControl OnHeaderFooter(HeaderFooter headerFooter) => throw new InvalidOperationException();

        public WorksheetReadControl OnSheetDataBegin() => Record("begin");

        public WorksheetReadControl OnScanResult(int maxRowIndex, int maxColumnIndex) => throw new InvalidOperationException();

        public WorksheetReadControl OnRow(int index, bool hidden, double? height) => Record("row:" + index);

        public WorksheetReadControl OnCell(in Cell cell) => Record("cell:" + cell.ColumnIndex + ":" + cell.Value);

        public WorksheetReadControl OnRowEnd()
        {
            Events.Add("end");
            return WorksheetReadControl.Pause;
        }

        private WorksheetReadControl Record(string value)
        {
            Events.Add(value);
            return pauseEveryEvent ? WorksheetReadControl.Pause : WorksheetReadControl.Continue;
        }
    }
}