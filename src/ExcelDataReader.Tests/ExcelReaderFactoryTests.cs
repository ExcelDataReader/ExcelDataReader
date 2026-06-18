namespace ExcelDataReader.Tests;

[TestFixture]
public class ExcelReaderFactoryTests
{
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
