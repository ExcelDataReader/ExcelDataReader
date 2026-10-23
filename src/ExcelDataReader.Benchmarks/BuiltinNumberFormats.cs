using System.Reflection;
using BenchmarkDotNet.Attributes;
using ExcelDataReader.Core.NumberFormat;
#if NET8_0_OR_GREATER
using System.Collections.Frozen;
#endif

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class BuiltinNumberFormats
{
    private Dictionary<int, NumberFormatString> _dictionary = null!;
    private NumberFormatString[] _array = null!;
    private int[] _indices = [];
#if NET8_0_OR_GREATER
    private FrozenDictionary<int, NumberFormatString> _frozen = null!;
#endif

    [Params("common", "all", "mixed")]
    public string Keys { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        var type = typeof(ExcelReaderFactory).Assembly.GetType("ExcelDataReader.Core.BuiltinNumberFormat")!;
        var get = type.GetMethod("GetBuiltinNumberFormat", BindingFlags.Public | BindingFlags.Static, null, [typeof(int)], null)!;
        _dictionary = [];
        for (int i = 0; i < 82; i++)
        {
            if (get.Invoke(null, [i]) is NumberFormatString value)
                _dictionary.Add(i, value);
        }

        _array = CreateArray();
#if NET8_0_OR_GREATER
        _frozen = CreateFrozen();
#endif
        int[] keys = Keys switch
        {
            "common" => [0, 0, 0, 0, 1, 2, 4, 9, 14, 14, 20, 22, 49],
            "all" => Enumerable.Range(0, 82).ToArray(),
            _ => [-1, 0, 2, 14, 23, 26, 49, 63, 66, 81, 82, 164, 65535, int.MaxValue],
        };
        _indices = Enumerable.Range(0, 1024).Select(i => keys[i % keys.Length]).ToArray();
        if (DictionaryLookup() != ArrayLookup())
            throw new InvalidOperationException("Lookup results differ.");
#if NET8_0_OR_GREATER
        if (DictionaryLookup() != FrozenLookup())
            throw new InvalidOperationException("Frozen lookup results differ.");
#endif
    }

    [Benchmark(Baseline = true, OperationsPerInvoke = 1024)]
    public int DictionaryLookup()
    {
        int count = 0;
        foreach (int index in _indices)
            count += _dictionary.TryGetValue(index, out var value) && value != null ? 1 : 0;
        return count;
    }

    [Benchmark(OperationsPerInvoke = 1024)]
    public int ArrayLookup()
    {
        int count = 0;
        foreach (int index in _indices)
            count += (uint)index < (uint)_array.Length && _array[index] != null ? 1 : 0;
        return count;
    }

#if NET8_0_OR_GREATER
    [Benchmark(OperationsPerInvoke = 1024)]
    public int FrozenLookup()
    {
        int count = 0;
        foreach (int index in _indices)
            count += _frozen.TryGetValue(index, out var value) && value != null ? 1 : 0;
        return count;
    }

    internal FrozenDictionary<int, NumberFormatString> CreateFrozen() => _dictionary.ToFrozenDictionary();
#endif

    internal NumberFormatString[] CreateArray()
    {
        var result = new NumberFormatString[82];
        foreach (var pair in _dictionary)
            result[pair.Key] = pair.Value;
        return result;
    }

    internal Dictionary<int, NumberFormatString> CreateDictionary() => new(_dictionary);
}