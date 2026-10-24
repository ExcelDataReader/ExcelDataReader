namespace ExcelDataReader.Tests;

[TestFixture]
public class ExcelReaderFactoryTests
{
    [Test]
    public void FactoriesRejectNullStreams(
        [Values("auto", "binary", "openxml", "csv")] string format,
        [Values(false, true)] bool leaveOpen)
    {
        var configuration = new ExcelReaderConfiguration { LeaveOpen = leaveOpen };
        var error = Assert.Throws<ArgumentNullException>(() =>
        {
            using var reader = format switch
            {
                "binary" => ExcelReaderFactory.CreateBinaryReader(null, configuration),
                "openxml" => ExcelReaderFactory.CreateOpenXmlReader(null, configuration),
                "csv" => ExcelReaderFactory.CreateCsvReader(null, configuration),
                _ => ExcelReaderFactory.CreateReader(null, configuration),
            };
        });
        Assert.That(error.ParamName, Is.EqualTo("fileStream"));
    }

    [Test]
    public void SpreadsheetXmlProbePreservesStreamFailure()
    {
        using var stream = new ProbeFailureStream();

        Assert.That(
            Assert.Throws<IOException>(() => ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration { LeaveOpen = true })),
            Is.SameAs(stream.Failure));
    }

    [TestCase("not a spreadsheet")]
    [TestCase("<Workbook invalid=\"&broken;\">")]
    public void UnrecognizedOrMalformedXmlStillReportsHeaderFailure(string text)
    {
        using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(text));

        Assert.Throws<Exceptions.HeaderException>(() => ExcelReaderFactory.CreateReader(stream));
    }

    [TestCase("10x10.xls")]
    [TestCase("UnicodeChars.xls")]
    [TestCase("biff3.xls")]
    [TestCase("as3xls_BIFF2.xls")]
    public void ProbeXls(string name)
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook(name));
        Assert.That(excelReader.GetType().Name, Is.EqualTo("ExcelBinaryReader"));
    }

    [TestCase("10x10.xlsx")]
    [TestCase("Open.xlsx")]
    [TestCase("Open.xlsb")]
    public void ProbeOpenXml(string name)
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook(name));
        Assert.That(excelReader.GetType().Name, Is.EqualTo("ExcelOpenXmlReader"));
    }

    [TestCase("SpreadsheetXml2003.xml")]
    [TestCase("SpreadsheetXml2003_LeadingWhitespace.xml")]
    public void ProbeSpreadsheetXml(string name)
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook(name));
        Assert.That(excelReader.GetType().Name, Is.EqualTo("ExcelSpreadsheetXmlReader"));
    }

    [Test]
    public void CreateReader_NonSeekableBinaryStream_Succeeds()
    {
        using var stream = Configuration.GetTestWorkbook("10x10.xls");
        using var nonSeekableStream = SeekErrorMemoryStream.CreateFromStream(stream);
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(nonSeekableStream);

        Assert.That(excelReader.GetType().Name, Is.EqualTo("ExcelBinaryReader"));
        Assert.That(excelReader.Read(), Is.True);
    }

    [Test]
    public void CreateBinaryReader_NonSeekableStream_Succeeds()
    {
        using var stream = Configuration.GetTestWorkbook("10x10.xls");
        using var nonSeekableStream = SeekErrorMemoryStream.CreateFromStream(stream);
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(nonSeekableStream);

        Assert.That(excelReader.Read(), Is.True);
    }

    [Test]
    public void CreateOpenXmlReader_NonSeekableStream_Succeeds()
    {
        using var stream = Configuration.GetTestWorkbook("10x10.xlsx");
        using var nonSeekableStream = SeekErrorMemoryStream.CreateFromStream(stream);
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(nonSeekableStream);

        Assert.That(excelReader.GetType().Name, Is.EqualTo("ExcelOpenXmlReader"));
        Assert.That(excelReader.Read(), Is.True);
    }

    [Test]
    public void CreateCsvReader_NonSeekableStream_Succeeds()
    {
        using var stream = Configuration.GetTestWorkbook(Path.Combine("csv", "MOCK_DATA.csv"));
        using var nonSeekableStream = SeekErrorMemoryStream.CreateFromStream(stream);
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateCsvReader(nonSeekableStream);

        Assert.That(excelReader.Read(), Is.True);
    }

    [Test]
    public void CreateReader_NonSeekableCopyFailure_DisposesSourceWhenLeaveOpenFalse()
    {
        var stream = new ThrowOnReadNonSeekableStream();

        Assert.Throws<IOException>(() => ExcelReaderFactory.CreateReader(stream));
        Assert.That(stream.IsDisposed, Is.True);
    }

    [Test]
    public void CreateReader_NonSeekableCopyFailure_DoesNotDisposeSourceWhenLeaveOpenTrue()
    {
        var stream = new ThrowOnReadNonSeekableStream();

        Assert.Throws<IOException>(() => ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration { LeaveOpen = true }));
        Assert.That(stream.IsDisposed, Is.False);
        stream.Dispose();
    }

    private sealed class ProbeFailureStream : MemoryStream
    {
        private bool _initialRead;

        public ProbeFailureStream()
            : base(System.Text.Encoding.UTF8.GetBytes("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"/>"))
        {
        }

        public IOException Failure { get; } = new("Input failed during format detection.");

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (_initialRead)
                throw Failure;
            _initialRead = true;
            return base.Read(buffer, offset, count);
        }

#if NET8_0_OR_GREATER
        public override int Read(Span<byte> buffer)
        {
            var bytes = new byte[buffer.Length];
            int count = Read(bytes, 0, bytes.Length);
            bytes.AsSpan(0, count).CopyTo(buffer);
            return count;
        }
#endif

        public override int ReadByte() => throw Failure;
    }

    private sealed class ThrowOnReadNonSeekableStream : Stream
    {
        public bool IsDisposed { get; private set; }

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count) => throw new IOException("Simulated read failure");

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            IsDisposed = true;
            base.Dispose(disposing);
        }
    }
}