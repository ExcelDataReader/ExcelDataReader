using System.Reflection;
using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class SharedStringConstruction
{
    private string _path;
    private object _configuration;
    private MethodInfo _create;

    [Params("Default", "SpillToDisk1", "SpillToDisk256")]
    public string Storage { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        _path = Environment.GetEnvironmentVariable("EDR_SST_INPUT")
            ?? throw new InvalidOperationException("Set EDR_SST_INPUT to an existing benchmark workbook.");
        string library = Environment.GetEnvironmentVariable("EDR_SST_LIBRARY");
        Assembly assembly = string.IsNullOrEmpty(library) ? typeof(ExcelReaderFactory).Assembly : Assembly.LoadFile(Path.GetFullPath(library));
        Type configurationType = assembly.GetType("ExcelDataReader.ExcelReaderConfiguration")!;
        _configuration = Activator.CreateInstance(configurationType)!;
        configurationType.GetProperty("SinglePassMode")!.SetValue(_configuration, true);
        configurationType.GetProperty("SharedStringStorageMode")!.SetValue(
            _configuration,
            Enum.Parse(assembly.GetType("ExcelDataReader.SharedStringStorageMode")!, Storage == "Default" ? "Default" : "SpillToDisk"));
        configurationType.GetProperty("SharedStringSpillThreshold")!.SetValue(_configuration, (Storage == "SpillToDisk1" ? 1L : 256L) * 1024 * 1024);
        configurationType.GetProperty("SharedStringTemporaryDirectory")!.SetValue(_configuration, Environment.CurrentDirectory);
        _create = assembly.GetType("ExcelDataReader.ExcelReaderFactory")!.GetMethod("CreateReader", [typeof(Stream), configurationType])!;
    }

    [Benchmark]
    public void Open()
    {
        using var input = File.OpenRead(_path);
        using var reader = (IDisposable)_create.Invoke(null, [input, _configuration])!;
    }
}
