using System.Text;
using ExcelDataReader.Core.BinaryFormat;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Tests;

public class XlsStringDataTests
{
    [TestCase(true)]
    [TestCase(false)]
    public void CompressedUnicodeOwnsDecodedCharacters(bool shortCount)
    {
        byte[] characters = [0, 65, 128, 233, 255];
        int offset = shortCount ? 6 : 7;
        var bytes = new byte[offset + characters.Length];
        bytes[2] = (byte)(bytes.Length - 4);
        bytes[4] = (byte)characters.Length;
        Array.Copy(characters, 0, bytes, offset, characters.Length);

        var value = XlsStringData.Read(new XlsBiffFrame(bytes), 4, shortCount, true);
        Array.Clear(bytes, 0, bytes.Length);

        Assert.That(value.GetValue(Encoding.Unicode), Is.EqualTo("\0A\u0080\u00e9\u00ff"));
    }

    [Test]
    public void CompressedUnicodeRejectsTruncatedCharacters()
    {
        byte[] bytes = [0, 0, 3, 0, 2, 0, 65];
        Assert.Throws<ExcelReaderException>(() => XlsStringData.Read(new XlsBiffFrame(bytes), 4, true, true));
    }
}