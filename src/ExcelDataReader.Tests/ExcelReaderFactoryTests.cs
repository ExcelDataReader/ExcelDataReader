namespace ExcelDataReader.Tests;

[TestFixture]
public class ExcelReaderFactoryTests
{
    [TestCase(1)]
    [TestCase(7)]
    public void RawXlsAcceptsShortReads(int chunkSize)
    {
        using var data = new MemoryStream();
        using (var writer = new BinaryWriter(data, System.Text.Encoding.Unicode, leaveOpen: true))
        {
            writer.Write((ushort)0x809);
            writer.Write((ushort)16);
            writer.Write((ushort)0x600);
            writer.Write((ushort)0x10);
            writer.Write(new byte[12]);
            writer.Write((ushort)0x203);
            writer.Write((ushort)14);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write(42.0);
            writer.Write((ushort)0xA);
            writer.Write((ushort)0);
        }

        using var stream = new ShortReadStream(data.ToArray(), chunkSize);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(42));
        Assert.That(reader.Read(), Is.False);
    }

    [Test]
    public void XlsTransportHandlesShortReadsAndRejectsTruncation(
        [Values(1, 7, 1024)] int chunkSize,
        [Values(20, 1024, 3000)] int payloadLength,
        [Values(false, true)] bool truncated)
    {
        using var data = new MemoryStream();
        byte[] payload = Enumerable.Range(0, payloadLength).Select(i => (byte)i).ToArray();
        using (var writer = new BinaryWriter(data, System.Text.Encoding.Unicode, leaveOpen: true))
        {
            writer.Write((ushort)0xFF);
            writer.Write((ushort)1018);
            writer.Write(new byte[1018]);
            writer.Write((ushort)0xFF);
            writer.Write((ushort)payloadLength);
            writer.Write(payload);
        }

        byte[] bytes = data.ToArray();
        if (truncated)
            Array.Resize(ref bytes, bytes.Length - 1);
        using var source = new ShortReadStream(bytes, chunkSize);
        using var transport = new Core.BinaryFormat.XlsBiffStream(source);
        Assert.That(transport.ReadFrame(out _), Is.True);
        if (truncated)
        {
            Assert.Throws<EndOfStreamException>(() => transport.ReadFrame(out _));
        }
        else
        {
            Assert.That(transport.ReadFrame(out var frame), Is.True);
            Assert.That(frame.Bytes.Skip(4).Take(payloadLength), Is.EqualTo(payload));
            Assert.That(transport.ReadFrame(out _), Is.False);
        }
    }

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

    private sealed class ShortReadStream(byte[] bytes, int chunkSize) : MemoryStream(bytes)
    {
        public override int Read(byte[] buffer, int offset, int count) => base.Read(buffer, offset, Math.Min(count, chunkSize));

#if NET8_0_OR_GREATER
        public override int Read(Span<byte> buffer) => base.Read(buffer[..Math.Min(buffer.Length, chunkSize)]);
#endif
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