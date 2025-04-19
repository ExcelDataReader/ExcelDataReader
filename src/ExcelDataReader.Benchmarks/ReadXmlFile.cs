using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class ReadXmlFile
{
    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    [Benchmark]
    public void ReadSingleFileXslx()
    {
        using var reader = ExcelReaderFactory.CreateReader(typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test10x10000.xlsx"));
        Read(reader);
    }

    [Benchmark]
    public void ReadSingleFileXslb()
    {
        using var reader = ExcelReaderFactory.CreateReader(typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test10x10000.xlsb"));
        Read(reader);
    }

    [Benchmark]
    public void ReadSingleFileXls()
    {
        using var reader = ExcelReaderFactory.CreateReader(typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test10x10000.xls"));
        Read(reader);
    }

    private static void Read(IExcelDataReader reader)
    {
        while (reader.Read())
        {            
        }        
    }
}