using BenchmarkDotNet.Attributes;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class ParseLongFormats
{
    private string _format = string.Empty;

    [Params("Digits", "Dates", "Directives", "Unicode")]
    public string Content { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        _format = Content switch
        {
            "Digits" => new string('0', 1024) + ".00",
            "Dates" => string.Concat(Enumerable.Repeat("yyyy-mm-dd hh:mm:ss.000 ", 32)),
            "Directives" => string.Concat(Enumerable.Repeat("[>=100][Red]", 64)) + "0.00",
            "Unicode" => string.Concat(Enumerable.Repeat("yyyy\"\u5E74\"m\"\u6708\"d\"\u65E5\"", 32)),
            _ => throw new InvalidOperationException(Content),
        };
    }

    [Benchmark]
    public NumberFormatString Parse() => new(_format);
}
