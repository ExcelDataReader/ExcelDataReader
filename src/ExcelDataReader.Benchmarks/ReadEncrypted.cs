using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class ReadEncrypted
{
    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    [Benchmark]
    public void Xlsx()
    {
        using var reader = ExcelReaderFactory.CreateReader(typeof(OpenXmlFile).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.Test_git_issue289.xlsx"));
        Read(reader);
    }

    private static void Read(IExcelDataReader reader)
    {
        while (reader.Read())
        {            
        }        
    }
}