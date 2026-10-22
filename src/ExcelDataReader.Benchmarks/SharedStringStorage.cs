#nullable enable

using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Reflection;
using System.Text;
#if NET8_0_OR_GREATER
using System.Text.Json;
#endif
using BenchmarkDotNet.Attributes;
using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class SharedStringStorage
{
    private string _path = string.Empty;

    [Params("xlsx", "xlsb", "xls")]
    public string Format { get; set; } = "xlsx";

    [Params("Sequential", "Permuted")]
    public string Pattern { get; set; } = "Sequential";

    [Params(1000000)]
    public int Count { get; set; }

    [Params(64)]
    public int Length { get; set; }

    [Params("Default", "SpillToDisk1", "SpillToDisk16", "SpillToDisk64", "SpillToDisk256")]
    public string Storage { get; set; } = "Default";

    private (long Cells, long Characters, ulong Checksum) ExpectedResult { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        _path = Path.Combine(Environment.CurrentDirectory, "ExcelDataReader-sst-" + Guid.NewGuid().ToString("N") + "." + Format);
        SharedStringWorkbook.Create(_path, Format, Count, Length, Pattern);
        var actual = Read(_path, Configuration("Default", false));
        ExpectedResult = Expected(Count, Length, Pattern);
        if (actual != ExpectedResult)
            throw new InvalidOperationException("Generated corpus failed output validation.");
    }

    [GlobalCleanup]
    public void Cleanup() => File.Delete(_path);

    [Benchmark]
    public long ReadAll()
    {
        var actual = Read(_path, Configuration(Storage, false));
        if (actual != ExpectedResult)
            throw new InvalidOperationException("Storage mode changed workbook output.");
        return actual.Cells;
    }

    internal static (long Cells, long Characters, ulong Checksum) Read(string path, ExcelReaderConfiguration configuration)
    {
        using var stream = File.OpenRead(path);
        using var reader = ExcelReaderFactory.CreateReader(stream, configuration);
        long cells = 0;
        long characters = 0;
        ulong checksum = 14695981039346656037;
        do
        {
            while (reader.Read())
            {
                for (int c = 0; c < reader.FieldCount; c++)
                {
                    if (reader.GetValue(c) is string value)
                    {
                        cells++;
                        characters += value.Length;
                        checksum = SharedStringWorkbook.Hash(checksum, value);
                    }
                }
            }
        }
        while (reader.NextResult());
        return (cells, characters, checksum);
    }

    internal static void RunScenario(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        if (args[0] == "--sst-generate")
        {
            string path = args[1];
            int count = int.Parse(args[2], CultureInfo.InvariantCulture);
            int length = int.Parse(args[3], CultureInfo.InvariantCulture);
            string pattern = args.Length > 4 ? args[4] : "Sequential";
            bool unicode = args.Length > 5 && bool.Parse(args[5]);
            SharedStringWorkbook.Create(path, Path.GetExtension(path).TrimStart('.'), count, length, pattern, unicode);
            var actual = Read(path, new ExcelReaderConfiguration());
            if (actual != Expected(count, length, pattern, unicode))
                throw new InvalidOperationException("Generated corpus failed output validation.");
            Print(new { Path = path, Seed = 741, Count = count, Length = length, Pattern = pattern, Unicode = unicode, Bytes = new FileInfo(path).Length, SstBytes = SstBytes(path), actual.Cells, actual.Characters, actual.Checksum });
            return;
        }

        string input = args[1];
        bool singlePass = args.Length > 2 && bool.Parse(args[2]);
        string storage = args.Length > 3 ? args[3] : "Default";
        var configuration = Configuration(storage, singlePass);
        using var process = Process.GetCurrentProcess();
        WarmUp();
        long baseline = GC.GetTotalMemory(true);
        long baselineWorkingSet = process.WorkingSet64;
        var timer = Stopwatch.StartNew();
        using var stream = File.OpenRead(input);
        using var reader = ExcelReaderFactory.CreateReader(stream, configuration);
        double constructionMs = timer.Elapsed.TotalMilliseconds;
        long retainedAfterOpen = GC.GetTotalMemory(true) - baseline;
        object? store = GetStore(reader);
        bool spilledAfterOpen = GetStatistic<bool>(store, "HasSpilled");
        long accountedBytes = GetStatistic<long>(store, "ResidentBytes");
        long diskBytes = GetStatistic<long>(store, "DiskBytes");
        long sampledHeapPeak = GC.GetTotalMemory(false);
        long? retainedAfterPartialRead = null;
        long cells = 0;
        long characters = 0;
        ulong checksum = 14695981039346656037;
        timer.Restart();
        do
        {
            while (reader.Read())
            {
                for (int c = 0; c < reader.FieldCount; c++)
                {
                    if (reader.GetValue(c) is string value)
                    {
                        cells++;
                        characters += value.Length;
                        checksum = SharedStringWorkbook.Hash(checksum, value);
                    }
                }

                if (cells % 4096 < 4)
                    sampledHeapPeak = Math.Max(sampledHeapPeak, GC.GetTotalMemory(false));
                if (retainedAfterPartialRead == null && cells >= 10000)
                    retainedAfterPartialRead = GC.GetTotalMemory(true) - baseline;
            }
        }
        while (reader.NextResult());
        double traversalMs = timer.Elapsed.TotalMilliseconds;
        long retainedAfterRead = GC.GetTotalMemory(true) - baseline;
        accountedBytes = Math.Max(accountedBytes, GetStatistic<long>(store, "ResidentBytes"));
        bool spilled = GetStatistic<bool>(store, "HasSpilled");
        diskBytes = GetStatistic<long>(store, "DiskBytes");
        process.Refresh();
        long peakWorkingSet = process.PeakWorkingSet64;
        reader.Dispose();
        store = null;
        long retainedAfterDispose = GC.GetTotalMemory(true) - baseline;
        Print(new { Path = input, Storage = storage, configuration.SinglePassMode, ServerGC = System.Runtime.GCSettings.IsServerGC, constructionMs, traversalMs, cells, characters, checksum, retainedAfterOpen, retainedAfterPartialRead, retainedAfterRead, retainedAfterDispose, baselineWorkingSet, peakWorkingSet, sampledHeapPeak, spilledAfterOpen, spilled, accountedBytes, diskBytes });
    }

    private static (long Cells, long Characters, ulong Checksum) Expected(int count, int length, string pattern, bool unicode = false)
    {
        int references = SharedStringWorkbook.References(count, pattern);
        ulong hash = 14695981039346656037;
        for (int i = 0; i < references; i++)
            hash = SharedStringWorkbook.Hash(hash, SharedStringWorkbook.Value(SharedStringWorkbook.Reference(i, count, pattern), length, unicode));
        return (references, (long)references * length, hash);
    }

    private static ExcelReaderConfiguration Configuration(string storage, bool singlePass) => new()
    {
        SinglePassMode = singlePass,
        SharedStringStorageMode = storage switch
        {
            "Default" => SharedStringStorageMode.Default,
            "SpillToDisk1" or "SpillToDisk16" or "SpillToDisk64" or "SpillToDisk256" => SharedStringStorageMode.SpillToDisk,
            _ => throw new ArgumentException("Unknown storage profile.", nameof(storage)),
        },
        SharedStringSpillThreshold = storage switch
        {
            "SpillToDisk1" => 1024L * 1024,
            "SpillToDisk16" => 16L * 1024 * 1024,
            "SpillToDisk256" => 256L * 1024 * 1024,
            _ => 64L * 1024 * 1024,
        },
        SharedStringTemporaryDirectory = Environment.CurrentDirectory,
    };

    private static object? GetStore(IExcelDataReader reader)
    {
        object? workbook = reader.GetType().GetProperty("Workbook", BindingFlags.Instance | BindingFlags.NonPublic)?.GetValue(reader);
        return workbook?.GetType().GetField("_stringStore", BindingFlags.Instance | BindingFlags.NonPublic)?.GetValue(workbook);
    }

    private static long SstBytes(string path)
    {
        using var input = File.OpenRead(path);
        if (Path.GetExtension(path) != ".xls")
        {
            using var archive = new ZipArchive(input, ZipArchiveMode.Read);
            return archive.Entries.Single(e => e.FullName.StartsWith("xl/sharedStrings.", StringComparison.Ordinal)).Length;
        }

        using var reader = new BinaryReader(input);
        long bytes = 0;
        while (input.Position < input.Length)
        {
            ushort id = reader.ReadUInt16();
            ushort size = reader.ReadUInt16();
            if (id == 0xA)
                break;
            if (id is 0xFC or 0x3C)
                bytes += 4L + size;
            input.Position += size;
        }

        return bytes;
    }

    private static T GetStatistic<T>(object? store, string name) =>
        store == null ? default! : (T)store.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(store)!;

    private static void WarmUp()
    {
        using var input = new MemoryStream(AllocationTestWorkbook.CreateXlsx(["<t>warmup</t>"], true));
        using var reader = ExcelReaderFactory.CreateReader(input);
        reader.Read();
        reader.GetString(0);
    }

    private static void Print(object value)
    {
#if NET8_0_OR_GREATER
        Console.WriteLine(JsonSerializer.Serialize(value));
#else
        Console.WriteLine(string.Join("; ", value.GetType().GetProperties().Select(p => p.Name + "=" + p.GetValue(value))));
#endif
    }
}
