using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class OpenXmlFile
{
    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    [Benchmark]
    public bool OpenSingleFileXslx()
    {
        using var reader = ExcelReaderFactory.CreateReader(typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test10x10000.xlsx"));
        return reader.Read();
    }

    [Benchmark]
    public bool OpenSingleFileXslb()
    {
        using var reader = ExcelReaderFactory.CreateReader(typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test10x10000.xlsb"));
        return reader.Read();
    }

    [Benchmark]
    public bool OpenSingleFileXls()
    {
        using var reader = ExcelReaderFactory.CreateReader(typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test10x10000.xls"));
        return reader.Read();
    }
}
