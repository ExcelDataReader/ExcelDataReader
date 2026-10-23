using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class BuiltinNumberFormatConstruction
{
    private readonly BuiltinNumberFormats _formats = new() { Keys = "common" };

    [GlobalSetup]
    public void Setup() => _formats.Setup();

    [Benchmark(Baseline = true)]
    public object Dictionary() => _formats.CreateDictionary();

    [Benchmark]
    public object Array() => _formats.CreateArray();

#if NET8_0_OR_GREATER
    [Benchmark]
    public object Frozen() => _formats.CreateFrozen();
#endif
}