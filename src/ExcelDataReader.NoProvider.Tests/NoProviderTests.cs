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
}
