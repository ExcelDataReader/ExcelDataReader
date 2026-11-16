using System.Runtime.Versioning;
using System.Text;
using ExcelDataReader.Tests;

namespace ExcelDataReader.NoProvider.Tests;

public class NoProviderTests
{
    [Test]
    public void UsesExpectedLibraryAsset()
    {
        string framework = typeof(ExcelReaderFactory).Assembly.GetCustomAttributes(typeof(TargetFrameworkAttribute), false)
            .Cast<TargetFrameworkAttribute>().Single().FrameworkName;
#if NET10_0
        Assert.That(framework, Is.EqualTo(".NETStandard,Version=v2.0"));
#else
        Assert.That(framework, Is.EqualTo(".NETCoreApp,Version=v10.0"));
#endif
    }

    [TestCase("10x10.xlsx", false, false)]
    [TestCase("10x10.xlsx", false, true)]
    [TestCase("10x10.xlsx", true, false)]
    [TestCase("10x10.xlsx", true, true)]
    [TestCase("10x10.xlsb", false, false)]
    [TestCase("10x10.xlsb", false, true)]
    [TestCase("10x10.xls", false, false)]
    [TestCase("10x10.xls", false, true)]
    [TestCase("10x10.xml", false, false)]
    [TestCase("10x10.xml", false, true)]
    public void ReadsUnicodeWorkbooksWithoutRegisteringProvider(string file, bool strict, bool singlePass)
    {
        AssertProviderNotRegistered();
        using var stream = TestWorkbook.Open(strict ? Path.Combine("strict", file) : file);
        using var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Read(), Is.True);
        AssertProviderNotRegistered();
    }

    [Test]
    public void ReadsUtf8CsvWithoutRegisteringProvider()
    {
        AssertProviderNotRegistered();
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("a,b\n"));
        using var reader = ExcelReaderFactory.CreateCsvReader(stream);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("a"));
        AssertProviderNotRegistered();
    }

    [Test]
    public void LegacyCsvFallbackWithoutRegisteringProvider()
    {
        AssertProviderNotRegistered();
        using var stream = new MemoryStream([0xE9, (byte)',', (byte)'b', (byte)'\n']);
#if NET10_0
        Assert.Throws<NotSupportedException>(() => ExcelReaderFactory.CreateCsvReader(stream));
#else
        using var reader = ExcelReaderFactory.CreateCsvReader(stream);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("\u00e9"));
#endif
        AssertProviderNotRegistered();
    }

    [Test]
    public void IncompleteUtf8CsvUsesFallbackWithoutRegisteringProvider()
    {
        AssertProviderNotRegistered();
        using var stream = new MemoryStream([0xe9]);
#if NET10_0
        Assert.Throws<NotSupportedException>(() => ExcelReaderFactory.CreateCsvReader(stream));
#else
        using var reader = ExcelReaderFactory.CreateCsvReader(stream);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("\u00e9"));
#endif
        AssertProviderNotRegistered();
    }

    [TestCase(false)]
    [TestCase(true)]
    public void ReadsBiff2WithDeclaredUtf8WithoutRegisteringProvider(bool singlePass)
    {
        AssertProviderNotRegistered();
        using var stream = CreateBiff2(65001);
        using var reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("A"));
        reader.Reset();
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("A"));
        AssertProviderNotRegistered();
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Biff2WithoutCodePageUsesConfiguredFallback(bool singlePass)
    {
        AssertProviderNotRegistered();
        using var stream = CreateBiff2(null);
        var configuration = new ExcelReaderConfiguration { SinglePassMode = singlePass, FallbackEncoding = Encoding.UTF8 };
        using var reader = ExcelReaderFactory.CreateBinaryReader(stream, configuration);
        configuration.FallbackEncoding = Encoding.Unicode;
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("A"));
        AssertProviderNotRegistered();
    }

    [TestCase(false)]
    [TestCase(true)]
    public void ReadsBiff5WithDeclaredUtf8WithoutRegisteringProvider(bool singlePass)
    {
        AssertProviderNotRegistered();
        using var stream = CreateBiff5();
        using var reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Name, Is.EqualTo("S"));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("A"));
        reader.Reset();
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("A"));
        AssertProviderNotRegistered();
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Biff2WithoutCodePageResolvesDefaultOnlyWhenNeeded(bool singlePass)
    {
        AssertProviderNotRegistered();
        using var stream = CreateBiff2(null);
        var configuration = new ExcelReaderConfiguration { SinglePassMode = singlePass };
        using var reader = ExcelReaderFactory.CreateBinaryReader(stream, configuration);
#if NET10_0
        Assert.Throws<NotSupportedException>(() => reader.Read());
#else
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("A"));
#endif
        AssertProviderNotRegistered();
    }

    [TestCase(false)]
    [TestCase(true)]
    public void ReadsBiff8FormulaStringsWithoutRegisteringProvider(bool singlePass)
    {
        AssertProviderNotRegistered();
        using var stream = TestWorkbook.Open("Hyperlinks.xls");
        using var reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration { SinglePassMode = singlePass });
        for (int row = 0; row <= 10; row++)
            Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("Formula label"));
        reader.Reset();
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("Website fragment"));
        AssertProviderNotRegistered();
    }

    [TestCase(0x0204, false)]
    [TestCase(0x0204, true)]
    [TestCase(0x00d6, false)]
    [TestCase(0x00d6, true)]
    public void ReadsBiff8InlineStringsWithoutRegisteringProvider(int recordId, bool singlePass)
    {
        AssertProviderNotRegistered();
        using var stream = CreateBiff8InlineString((ushort)recordId);
        using var reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("\u03bb"));
        reader.Reset();
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("\u03bb"));
        AssertProviderNotRegistered();
    }

    [TestCase("url", false)]
    [TestCase("url", true)]
    [TestCase("unicodeFile", false)]
    [TestCase("unicodeFile", true)]
    public void ReadsBiff8UnicodeHyperlinksWithoutRegisteringProvider(string kind, bool singlePass)
    {
        AssertProviderNotRegistered();
        using var stream = CreateBiff8Hyperlink(kind);
        using var reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetHyperlink(0)?.Address, Is.EqualTo(kind == "url" ? "https://example.com/" : "\u5831\u544a.xlsx"));
        AssertProviderNotRegistered();
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Biff8AnsiHyperlinkStillRequiresCodePage(bool singlePass)
    {
        AssertProviderNotRegistered();
        using var stream = CreateBiff8Hyperlink("ansiFile");
        using var reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Read(), Is.True);
#if NET10_0
        Assert.Throws<NotSupportedException>(() => reader.GetHyperlink(0));
#else
        Assert.That(reader.GetHyperlink(0)?.Address, Is.EqualTo("caf\u00e9.xls"));
#endif
        AssertProviderNotRegistered();
    }

    private static void AssertProviderNotRegistered()
    {
        Assert.Throws<NotSupportedException>(() => Encoding.GetEncoding(1252));
    }

    private static MemoryStream CreateBiff2(ushort? codePage)
    {
        var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true))
        {
            writer.Write((ushort)0x09);
            writer.Write((ushort)4);
            writer.Write((ushort)2);
            writer.Write((ushort)0x10);
            if (codePage.HasValue)
            {
                writer.Write((ushort)0x42);
                writer.Write((ushort)2);
                writer.Write(codePage.Value);
            }

            writer.Write((ushort)0x04);
            writer.Write((ushort)9);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write(new byte[3]);
            writer.Write((byte)1);
            writer.Write((byte)'A');
            writer.Write((ushort)0x0A);
            writer.Write((ushort)0);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateBiff5()
    {
        var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true))
        {
            writer.Write((ushort)0x0809);
            writer.Write((ushort)4);
            writer.Write((ushort)0x0500);
            writer.Write((ushort)0x0005);
            writer.Write((ushort)0x0085);
            writer.Write((ushort)8);
            writer.Write(30U);
            writer.Write((ushort)0);
            writer.Write((byte)1);
            writer.Write((byte)'S');
            writer.Write((ushort)0x0042);
            writer.Write((ushort)2);
            writer.Write((ushort)65001);
            writer.Write((ushort)0x000A);
            writer.Write((ushort)0);
            writer.Write((ushort)0x0809);
            writer.Write((ushort)4);
            writer.Write((ushort)0x0500);
            writer.Write((ushort)0x0010);
            writer.Write((ushort)0x0204);
            writer.Write((ushort)9);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)1);
            writer.Write((byte)'A');
            writer.Write((ushort)0x000A);
            writer.Write((ushort)0);
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateBiff8InlineString(ushort recordId)
    {
        var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true))
        {
            BiffRecord(writer, 0x0809, data =>
            {
                data.Write((ushort)0x0600);
                data.Write((ushort)5);
            });
            BiffRecord(writer, 0x0085, data =>
            {
                data.Write(42U);
                data.Write((ushort)0);
                data.Write((byte)1);
                data.Write((byte)0);
                data.Write((byte)'S');
            });
            BiffRecord(writer, 0x0031, data => data.Write(new byte[13]));
            BiffRecord(writer, 0x000a, _ => { });
            BiffRecord(writer, 0x0809, data =>
            {
                data.Write((ushort)0x0600);
                data.Write((ushort)0x10);
            });
            BiffRecord(writer, recordId, data =>
            {
                data.Write((ushort)0);
                data.Write((ushort)0);
                data.Write((ushort)0);
                data.Write((ushort)1);
                data.Write((byte)1);
                data.Write((ushort)0x03bb);
            });
            BiffRecord(writer, 0x000a, _ => { });
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateBiff8Hyperlink(string kind)
    {
        var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true))
        {
            BiffRecord(writer, 0x0809, data =>
            {
                data.Write((ushort)0x0600);
                data.Write((ushort)0x10);
            });
            BiffRecord(writer, 0x0203, data =>
            {
                data.Write((ushort)0);
                data.Write((ushort)0);
                data.Write((ushort)0);
                data.Write(42.0);
            });
            BiffRecord(writer, 0x01b8, data =>
            {
                data.Write(new byte[8]);
                data.Write(new Guid("79EAC9D0-BAF9-11CE-8C82-00AA004BA90B").ToByteArray());
                data.Write(2);
                data.Write(1);
                data.Write(new Guid(kind == "url" ? "79EAC9E0-BAF9-11CE-8C82-00AA004BA90B" : "00000303-0000-0000-C000-000000000046").ToByteArray());
                if (kind == "url")
                {
                    byte[] url = Encoding.Unicode.GetBytes("https://example.com/\0");
                    data.Write(url.Length);
                    data.Write(url);
                }
                else
                {
                    byte[] ansi = [(byte)'c', (byte)'a', (byte)'f', 0xe9, (byte)'.', (byte)'x', (byte)'l', (byte)'s', 0];
                    data.Write((ushort)0);
                    data.Write(ansi.Length);
                    data.Write(ansi);
                    data.Write(ushort.MaxValue);
                    data.Write((ushort)0xdead);
                    data.Write(new byte[20]);
                    byte[] unicode = kind == "unicodeFile" ? Encoding.Unicode.GetBytes("\u5831\u544a.xlsx") : [];
                    data.Write(unicode.Length == 0 ? 0 : unicode.Length + 6);
                    if (unicode.Length != 0)
                    {
                        data.Write(unicode.Length);
                        data.Write((ushort)3);
                        data.Write(unicode);
                    }
                }
            });
            BiffRecord(writer, 0x000a, _ => { });
        }

        stream.Position = 0;
        return stream;
    }

    private static void BiffRecord(BinaryWriter writer, ushort id, Action<BinaryWriter> write)
    {
        using var payload = new MemoryStream();
        using (var data = new BinaryWriter(payload, Encoding.UTF8, leaveOpen: true))
            write(data);
        writer.Write(id);
        writer.Write((ushort)payload.Length);
        writer.Write(payload.ToArray());
    }
}
