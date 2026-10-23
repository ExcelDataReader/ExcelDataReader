using BenchmarkDotNet.Attributes;
#if NET8_0_OR_GREATER
using System.Collections.Frozen;
#endif

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class CustomNumberFormats
{
    private Dictionary<int, object> _dictionary = null!;
    private int[] _indices = [];
#if NET8_0_OR_GREATER
    private FrozenDictionary<int, object> _frozen = null!;
#endif

    [Params(0, 1, 16, 256)]
    public int Count { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _dictionary = [];
        for (int i = 0; i < Count; i++)
            _dictionary.Add(164 + i, new object());
#if NET8_0_OR_GREATER
        _frozen = _dictionary.ToFrozenDictionary();
#endif
        _indices = Enumerable.Range(0, 1024).Select(i => i % 2 == 0 ? i % 82 : 164 + i % Math.Max(1, Count)).ToArray();
    }

    [Benchmark(Baseline = true, OperationsPerInvoke = 1024)]
    public int DictionaryLookup()
    {
        int count = 0;
        foreach (int index in _indices)
            count += _dictionary.TryGetValue(index, out _) ? 1 : 0;
        return count;
    }

#if NET8_0_OR_GREATER
    [Benchmark(OperationsPerInvoke = 1024)]
    public int FrozenLookup()
    {
        int count = 0;
        foreach (int index in _indices)
            count += _frozen.TryGetValue(index, out _) ? 1 : 0;
        return count;
    }
#endif
}