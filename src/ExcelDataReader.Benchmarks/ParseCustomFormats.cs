using BenchmarkDotNet.Attributes;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class ParseCustomFormats
{
    [Params("0", "0.00", "#,##0.00;[Red]-#,##0.00", "[>=100]0.00;[Blue][h]:mm:ss", "yyyy-mm-dd hh:mm:ss.000", "0.00E+00", "0.00e-00", "# ?/?")]
    public string Format { get; set; } = string.Empty;

    [Benchmark]
    public NumberFormatString Parse() => new(Format);
}
