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

    private static void AssertProviderNotRegistered()
    {
        Assert.Throws<NotSupportedException>(() => Encoding.GetEncoding(1252));
    }
}
