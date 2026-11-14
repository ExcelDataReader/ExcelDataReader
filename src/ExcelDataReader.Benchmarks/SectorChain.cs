using System.Collections;
using System.Reflection;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class SectorChain
{
    private const uint EndOfChain = 0xfffffffe;
    private const uint FreeSpace = 0xffffffff;
    private List<uint> _table = [];
    private uint _first;
    private Func<uint, List<uint>, List<uint>> _current = null!;

    [Params("Small16", "Sparse16", "Medium1024", "Large16384", "Fragmented16384", "10x10000.xls", "BigFormatted.xls", "Issue184_FATSectors.xls")]
    public string Scenario { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        var type = typeof(ExcelReaderFactory).Assembly.GetType("ExcelDataReader.Core.CompoundFormat.CompoundDocument")!;
        _current = (Func<uint, List<uint>, List<uint>>)type.GetMethod("GetSectorChain", BindingFlags.Static | BindingFlags.NonPublic)!
            .CreateDelegate(typeof(Func<uint, List<uint>, List<uint>>));

        if (Scenario.EndsWith(".xls", StringComparison.Ordinal))
        {
            using var stream = typeof(SectorChain).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks." + Scenario)
                ?? throw new InvalidOperationException("Missing workbook " + Scenario);
            object document = Activator.CreateInstance(type, stream)!;
            _table = (List<uint>)type.GetProperty("SectorTable", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(document)!;
            var entries = (IEnumerable)type.GetProperty("Entries", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(document)!;
            bool found = false;
            foreach (object entry in entries)
            {
                var entryType = entry.GetType();
                string name = (string)entryType.GetProperty("EntryName")!.GetValue(entry)!;
                if (name is not ("Workbook" or "Book"))
                    continue;
                if ((bool)entryType.GetProperty("IsEntryMiniStream")!.GetValue(entry)!)
                    throw new InvalidOperationException("Workbook uses the mini sector table.");
                _first = (uint)entryType.GetProperty("StreamFirstSector")!.GetValue(entry)!;
                found = true;
                break;
            }

            if (!found)
                throw new InvalidOperationException("Workbook stream not found.");
        }
        else
        {
            int count = Scenario is "Small16" or "Sparse16" ? 16 : Scenario == "Medium1024" ? 1024 : 16384;
            _table = Enumerable.Repeat(EndOfChain, Scenario == "Sparse16" ? 30030 : count).ToList();
            uint sector = 0;
            _first = sector;
            for (int i = 1; i < count; i++)
            {
                uint next = Scenario == "Fragmented16384" ? (uint)((i * 7411) % count) : (uint)i;
                _table[(int)sector] = next;
                sector = next;
            }
        }

        var expected = Original();
        if (!expected.SequenceEqual(Current()) || !expected.SequenceEqual(Hybrid()) ||
            !expected.SequenceEqual(HashSet()) || !expected.SequenceEqual(BitSet()))
            throw new InvalidOperationException("Sector chain implementations disagree.");
        Console.WriteLine($"Scenario {Scenario}: {expected.Count} sectors, {_table.Count} FAT entries.");

        var savedTable = _table;
        uint savedFirst = _first;
        foreach (int length in new[] { 1, 64 })
        {
            _table = Enumerable.Range(1, length).Select(i => (uint)i).ToList();
            _table[length - 1] = length == 1 ? 0u : 7u;
            _first = 0;
            CheckCycle(Original);
            CheckCycle(Current);
            CheckCycle(Hybrid);
            CheckCycle(HashSet);
            CheckCycle(BitSet);
        }

        _table = savedTable;
        _first = savedFirst;
    }

    [Benchmark(Baseline = true)]
    public List<uint> Original()
    {
        List<uint> chain = [];
        uint sector = _first;
        while (sector != EndOfChain)
        {
            chain.Add(sector);
            sector = Next(sector);
            if (chain.Contains(sector))
                throw new InvalidOperationException("Cyclic sector chain");
        }

        Trim(chain);
        return chain;
    }

    [Benchmark]
    public List<uint> Current() => _current(_first, _table);

    [Benchmark]
    public List<uint> Hybrid()
    {
        const int threshold = 32;
        List<uint> chain = [];
        HashSet<uint> visited = null;
        uint sector = _first;
        while (sector != EndOfChain)
        {
            if (visited == null && chain.Count < threshold)
            {
                if (chain.Contains(sector))
                    throw new InvalidOperationException("Cyclic sector chain");
            }
            else
            {
                visited ??= new HashSet<uint>(chain);
                if (!visited.Add(sector))
                    throw new InvalidOperationException("Cyclic sector chain");
            }

            chain.Add(sector);
            sector = Next(sector);
        }

        Trim(chain);
        return chain;
    }

    [Benchmark]
    public List<uint> HashSet()
    {
        List<uint> chain = [];
        var visited = new HashSet<uint>();
        uint sector = _first;
        while (sector != EndOfChain)
        {
            if (!visited.Add(sector))
                throw new InvalidOperationException("Cyclic sector chain");
            chain.Add(sector);
            sector = Next(sector);
        }

        Trim(chain);
        return chain;
    }

    [Benchmark]
    public List<uint> BitSet()
    {
        List<uint> chain = [];
        var visited = new uint[(int)((_table.Count + 31L) / 32)];
        uint sector = _first;
        while (sector != EndOfChain)
        {
            if (sector < _table.Count)
            {
                int index = (int)(sector >> 5);
                uint bit = 1u << (int)(sector & 31);
                if ((visited[index] & bit) != 0)
                    throw new InvalidOperationException("Cyclic sector chain");
                visited[index] |= bit;
            }

            chain.Add(sector);
            sector = Next(sector);
        }

        Trim(chain);
        return chain;
    }

    private static void CheckCycle(Func<List<uint>> traverse)
    {
        try
        {
            traverse();
        }
        catch (InvalidOperationException)
        {
            return;
        }
        catch (Exception ex) when (ex.GetType().FullName == "ExcelDataReader.Exceptions.CompoundDocumentException")
        {
            return;
        }

        throw new InvalidOperationException("Sector chain cycle was not detected.");
    }

    private static void Trim(List<uint> chain)
    {
        while (chain.Count > 0 && chain[chain.Count - 1] == FreeSpace)
            chain.RemoveAt(chain.Count - 1);
    }

    private uint Next(uint sector) => sector < _table.Count ? _table[(int)sector] : EndOfChain;
}
