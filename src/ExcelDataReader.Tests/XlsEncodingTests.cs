using ExcelDataReader.Core;
using ExcelDataReader.Core.BinaryFormat;

namespace ExcelDataReader.Tests;

public class XlsEncodingTests
{
    [TestCase((ushort)1251, 1251)]
    [TestCase((ushort)932, 932)]
    [TestCase((ushort)10000, 10000)]
    [TestCase((ushort)65001, 65001)]
    public void EncodingHelperResolvesRepresentativeCodePages(ushort codePage, int expectedCodePage)
    {
        Assert.That(EncodingHelper.GetEncoding(codePage)!.CodePage, Is.EqualTo(expectedCodePage));
    }

    [Test]
    public void EncodingHelperReturnsNullForUnknownCodePage()
    {
        Assert.That(EncodingHelper.GetEncoding(9999), Is.Null);
    }

    [TestCase(204, 1251)]
    [TestCase(128, 932)]
    [TestCase(77, 10000)]
    [TestCase(255, 850)]
    [TestCase(99, 1252)]
    public void XlsFontCharsetSelectsItsByteStringEncoding(byte charset, int expectedCodePage)
    {
        byte[] record = new byte[17];
        record[0] = 0x31;
        record[2] = 13;
        record[16] = charset;
        var font = new XlsFontInfo(new XlsBiffFrame(record), 5);

        Assert.That(font.ByteStringEncoding.CodePage, Is.EqualTo(expectedCodePage));
    }

    [Test]
    public void XlsFontBeforeBiff5UsesDefaultByteStringEncoding()
    {
        byte[] record = new byte[17];
        record[0] = 0x31;
        record[2] = 13;
        record[16] = 204;
        var font = new XlsFontInfo(new XlsBiffFrame(record), 4);

        Assert.That(font.ByteStringEncoding.CodePage, Is.EqualTo(1252));
    }
}