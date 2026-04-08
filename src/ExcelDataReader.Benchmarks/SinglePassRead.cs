using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class SinglePassRead
{
    private static readonly ExcelReaderConfiguration MultiPassConfig = new();
    private static readonly ExcelReaderConfiguration SinglePassConfig = new() { SinglePassMode = true };

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    [Benchmark]
    public void ReadMultiPassXlsx()
    {
        using var reader = ExcelReaderFactory.CreateReader(
            typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test10x10000.xlsx"),
            MultiPassConfig);
        Read(reader);
    }

    [Benchmark]
    public void ReadSinglePassXlsx()
    {
        using var reader = ExcelReaderFactory.CreateReader(
            typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test10x10000.xlsx"),
            SinglePassConfig);
        Read(reader);
    }

    [Benchmark]
    public void ReadMultiPassXlsb()
    {
        using var reader = ExcelReaderFactory.CreateReader(
            typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test10x10000.xlsb"),
            MultiPassConfig);
        Read(reader);
    }

    [Benchmark]
    public void ReadSinglePassXlsb()
    {
        using var reader = ExcelReaderFactory.CreateReader(
            typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test10x10000.xlsb"),
            SinglePassConfig);
        Read(reader);
    }

    [Benchmark]
    public void ReadMultiPassXls()
    {
        using var reader = ExcelReaderFactory.CreateReader(
            typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test10x10000.xls"),
            MultiPassConfig);
        Read(reader);
    }

    [Benchmark]
    public void ReadSinglePassXls()
    {
        using var reader = ExcelReaderFactory.CreateReader(
            typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test10x10000.xls"),
            SinglePassConfig);
        Read(reader);
    }

    private static void Read(IExcelDataReader reader)
    {
        while (reader.Read())
        {
        }
    }
}
