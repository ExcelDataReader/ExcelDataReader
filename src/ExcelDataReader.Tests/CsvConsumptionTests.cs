using System.Text;
using ExcelDataReader.Core;
using ExcelDataReader.Core.CsvFormat;

namespace ExcelDataReader.Tests;

public class CsvConsumptionTests
{
    [Test]
    public void Issue687DoublePipeDelimiterPreservesEmptyColumns()
    {
        using var source = Configuration.GetTestWorkbook(Path.Combine("csv", "Issue687.csv"));
        using var reader = ExcelReaderFactory.CreateCsvReader(source, new ExcelReaderConfiguration { AutodetectSeparators = ["||"] });
        Assert.That(reader.FieldCount, Is.EqualTo(3));
        Assert.That(reader.RowCount, Is.EqualTo(4));
        string[][] expected =
        [
            ["field1", "field2", "field3"],
            ["value1", "value2", "value3"],
            [string.Empty, string.Empty, "value6"],
            ["value7", string.Empty, string.Empty],
        ];
        foreach (var row in expected)
        {
            Assert.That(reader.Read(), Is.True);
            for (int column = 0; column < row.Length; column++)
                Assert.That(reader.GetString(column), Is.EqualTo(row[column]));
        }

        Assert.That(reader.Read(), Is.False);
    }

    [TestCase("||", "||a||||\"b||c\"||d|", new[] { "", "a", "", "b||c", "d|" })]
    [TestCase("aab", "xaaabyaaba", new[] { "xa", "y", "a" })]
    [TestCase("abab", "ababaXababab", new[] { "", "aX", "ab" })]
    [TestCase(" | ", " | a | \"b | c\" | last ", new[] { "", "a", "b | c", "last" })]
    [TestCase("終é", "a終é\"b終éc\"終é終", new[] { "a", "b終éc", "終" })]
    [TestCase("\t|", "\ta\t| \tb \t|last\t", new[] { "a", "b", "last" })]
    public void MultiCharacterSeparatorsSurviveBufferSplitsAndCallbackPauses(string separator, string text, string[] fields)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        foreach (int bufferSize in new[] { 1, 2, 7, 1024 })
        {
            foreach (var mode in new[] { WorksheetReadMode.Materialize, WorksheetReadMode.Scan })
            {
                var parser = new CsvParser(separator, Encoding.UTF8, '"', mode: mode);
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

                var expected = new List<string> { "row:0" };
                if (mode == WorksheetReadMode.Materialize)
                    expected.AddRange(fields.Select((field, index) => $"field:{index}:{field}"));
                expected.Add($"end:{fields.Length}:False:True");
                Assert.That(handler.Events, Is.EqualTo(expected), $"{separator}, buffer {bufferSize}, {mode}");
            }
        }
    }

    [TestCase(0)]
    [TestCase(1)]
    public void MultiCharacterCandidatesPreserveAnalysisResetAndQuotedGrammar(int analyzeRows)
    {
        const string text = "a||\"b||c\"||\r\n\"escaped\\\"quote\"||\"d\"\"e\"||f\r\nlast|";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(text));
        using var reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration
        {
            AutodetectSeparators = [",", "||"],
            EscapeChar = '\\',
            AnalyzeInitialCsvRows = analyzeRows,
        });
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.FieldCount, Is.EqualTo(3));
            if (analyzeRows == 0)
                Assert.That(reader.RowCount, Is.EqualTo(3));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo("a"));
            Assert.That(reader.GetString(1), Is.EqualTo("b||c"));
            Assert.That(reader.GetString(2), Is.Empty);
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo("escaped\"quote"));
            Assert.That(reader.GetString(1), Is.EqualTo("d\"e"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo("last|"));
            Assert.That(reader.Read(), Is.False);
            reader.Reset();
        }
    }

    [TestCase(null)]
    [TestCase("")]
    [TestCase("a\r")]
    [TestCase("\nb")]
    [TestCase("a\"b")]
    public void InvalidSeparatorCandidatesAreRejected(string separator)
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("a,b"));
        Assert.Throws<ArgumentException>(() => ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration
        {
            AutodetectSeparators = [separator],
        }));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void LongOverlappingSeparatorsMatchAnIndependentSplit(bool trim)
    {
        var random = new Random(687);
        foreach (string separator in new[] { "aa", "aab", "abab", " | ", "\t|", new string('a', 2050) + "b" })
        {
            for (int sample = 0; sample < 12; sample++)
            {
                var text = new StringBuilder();
                for (int i = 0; i < 40; i++)
                    text.Append(random.Next(3) == 0 ? separator : "aaab \t|x"[random.Next(8)].ToString());
                text.Append('x');
                var expected = text.ToString().Split([separator], StringSplitOptions.None).Select(field => trim ? field.Trim(' ', '\t') : field);
                byte[] bytes = Encoding.UTF8.GetBytes(text.ToString());
                var parser = new CsvParser(separator, Encoding.UTF8, trimWhiteSpace: trim);
                var handler = new ParserHandler(pause: true);
                int offset = 0;
                int count = bytes.Length;
                while (parser.ParseBuffer(bytes, ref offset, ref count, ref handler) == WorksheetReadStatus.Paused)
                {
                }

                while (parser.Flush(ref handler) == WorksheetReadStatus.Paused)
                {
                }

                Assert.That(
                    handler.Events.Where(e => e.StartsWith("field:", StringComparison.Ordinal)),
                    Is.EqualTo(expected.Select((field, index) => $"field:{index}:{field}")),
                    $"{separator.Length}, sample {sample}");
            }
        }
    }

    [TestCase("||x||\r\np|\nq||", new[] { "", "x", "", "p|", "q", "" })]
    [TestCase(" || x || \n || ", new[] { "", "x", "", "", "" })]
    public void MultiCharacterSeparatorsPreserveEmptyFieldsAndLineBoundaries(string text, string[] fields)
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(text));
        using var reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration { AutodetectSeparators = ["||"] });
        var actual = new List<string>();
        while (reader.Read())
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (!reader.IsDBNull(i))
                    actual.Add(reader.GetString(i));
            }
        }

        Assert.That(actual, Is.EqualTo(fields));
    }

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
                var parser = new CsvParser(separator.ToString(), Encoding.UTF8, quote, trim, escape, mode);
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
        var parser = new CsvParser(",", Encoding.UTF8, '"');
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
        var parser = new CsvParser(",", Encoding.UTF8, '"', escapeChar: '\\');
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
            var parser = new CsvParser(",", Encoding.UTF8, '"', mode: mode);
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
        CsvAnalyzer.Analyze(stream, [","], Encoding.UTF8, 0, '"', true, null, out int fieldCount, out _, out _, out _, out int rowCount);
        Assert.That(fieldCount, Is.EqualTo(fields));
        Assert.That(rowCount, Is.EqualTo(rows));
    }

    [Test]
    public void AnalysisPreservesSamplingAndSeparatorTieBreaking()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("a,b\nc,d\ne,f\ng,h\nlater,has,many,fields\n"));
        CsvAnalyzer.Analyze(stream, [",", ";"], Encoding.UTF8, 1, '"', true, null, out int fields, out string separator, out _, out _, out int rows);
        Assert.That(fields, Is.EqualTo(2));
        Assert.That(separator, Is.EqualTo(","));
        Assert.That(rows, Is.EqualTo(-1));

        stream.Position = 0;
        CsvAnalyzer.Analyze(stream, [",", ";"], Encoding.UTF8, 0, '"', true, null, out fields, out separator, out _, out _, out rows);
        Assert.That(fields, Is.EqualTo(4));
        Assert.That(rows, Is.EqualTo(5));

        using var tied = new MemoryStream(Encoding.UTF8.GetBytes("a,b;c;d\n"));
        CsvAnalyzer.Analyze(tied, [",", ";"], Encoding.UTF8, 0, '"', true, null, out fields, out separator, out _, out _, out rows);
        Assert.That(separator, Is.EqualTo(";"));
        Assert.That(fields, Is.EqualTo(3));
    }

    [TestCase("utf8", ",")]
    [TestCase("utf16", ",")]
    [TestCase("utf16be", ",")]
    [TestCase("utf8", "終é")]
    [TestCase("utf16", "終é")]
    [TestCase("utf16be", "終é")]
    public void WorksheetPausesBeforeCellsAndKeepsDecodedBuffersAndBom(string encodingName, string separator)
    {
        Encoding encoding = encodingName switch
        {
            "utf16" => Encoding.Unicode,
            "utf16be" => Encoding.BigEndianUnicode,
            _ => new UTF8Encoding(true),
        };
        string value = new string('x', 1022) + "é終";
        byte[] bytes = encoding.GetPreamble().Concat(encoding.GetBytes(value + separator + "b\r\nlast" + separator)).ToArray();
        using var stream = new MemoryStream(bytes);
        var worksheet = new CsvWorksheet(stream, Encoding.UTF8, [separator], 0, '"');
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
        var worksheet = new CsvWorksheet(stream, Encoding.GetEncoding(1252), [","], 0);
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
        var worksheet = new CsvWorksheet(stream, Encoding.UTF8, [","], 0);
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