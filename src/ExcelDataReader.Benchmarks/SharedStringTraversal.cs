using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class SharedStringTraversal
{
    private string _path;
    private (long Cells, long Characters, ulong Checksum) _expected;

    [Params("Default", "SpillToDisk1", "SpillToDisk256")]
    public string Storage { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        _path = Environment.GetEnvironmentVariable("EDR_SST_INPUT")
            ?? throw new InvalidOperationException("Set EDR_SST_INPUT to an existing benchmark workbook.");
        _expected = SharedStringStorage.Read(_path, Configuration("Default"));
    }

    [Benchmark]
    public long ReadAll()
    {
        var actual = SharedStringStorage.Read(_path, Configuration(Storage));
        if (actual != _expected)
            throw new InvalidOperationException("Storage mode changed workbook output.");
        return actual.Cells;
    }

    private static ExcelReaderConfiguration Configuration(string storage) => new()
    {
        SinglePassMode = true,
        SharedStringStorageMode = storage == "Default" ? SharedStringStorageMode.Default : SharedStringStorageMode.SpillToDisk,
        SharedStringSpillThreshold = (storage == "SpillToDisk1" ? 1L : 256L) * 1024 * 1024,
        SharedStringTemporaryDirectory = Environment.CurrentDirectory,
    };
}
