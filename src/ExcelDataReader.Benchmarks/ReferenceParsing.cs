using System.Globalization;
using System.Reflection;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class ReferenceParsing
{
    private string[] _references = [];
    private ParseReference _parse = null!;

#if NET8_0_OR_GREATER
    private delegate bool ParseReference(ReadOnlySpan<char> value, out int column, out int row);
#else
    private delegate bool ParseReference(string value, out int column, out int row);
#endif

    [Params("Uppercase", "Lowercase", "Invalid")]
    public string Input { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        var method = typeof(ExcelReaderFactory).Assembly.GetType("ExcelDataReader.Core.ReferenceHelper")!
            .GetMethod("ParseReference", BindingFlags.Public | BindingFlags.Static)!;
        _parse = (ParseReference)Delegate.CreateDelegate(typeof(ParseReference), method);
        string[] columns = Input switch
        {
            "Uppercase" => ["A", "J", "AZ", "XFD"],
            "Lowercase" => ["a", "j", "az", "xfd"],
            _ => [string.Empty, "$A", "A!", "AZ"],
        };
        _references = Enumerable.Range(0, 1024)
            .Select(i => columns[i % columns.Length] + (i + 1).ToString(CultureInfo.InvariantCulture) + (Input == "Invalid" ? "!" : string.Empty))
            .ToArray();
    }

    [Benchmark(OperationsPerInvoke = 1024)]
    public int Parse()
    {
        int result = 0;
        foreach (string reference in _references)
        {
            if (_parse(reference, out int column, out int row))
                result += column + row;
        }

        return result;
    }
}