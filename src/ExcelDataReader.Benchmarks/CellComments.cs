#nullable enable

using System.Text;
using BenchmarkDotNet.Attributes;
using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class CellComments
{
    private byte[] _data = [];
    private IExcelDataReader _cachedReader = null!;
    private int _column;

    [Params("xls", "xlsx", "xlsb", "xml")]
    public string Format { get; set; } = string.Empty;

    [Params(false, true)]
    public bool HasComments { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        if (!HasComments && Format == "xml")
        {
            _data = AllocationTestWorkbook.CreateSpreadsheetXml(Enumerable.Repeat("<Data ss:Type=\"String\">value</Data>", 10000).ToArray());
        }
        else
        {
            string name = HasComments ? "OldIssue11573_BlankValues" : "10x10000";
            using var source = typeof(CellComments).Assembly.GetManifestResourceStream($"ExcelDataReader.Benchmarks.{name}.{Format}")!;
            using var buffer = new MemoryStream();
            source.CopyTo(buffer);
            _data = buffer.ToArray();
        }

        _cachedReader = Open();
        _cachedReader.Read();
        if (HasComments)
            _cachedReader.Read();
        _column = HasComments ? 2 : 0;
        _ = _cachedReader.GetComment(_column);
    }

    [GlobalCleanup]
    public void Cleanup() => _cachedReader.Dispose();

    [Benchmark(Baseline = true)]
    public int ReadWithoutComments()
    {
        using var reader = Open();
        int rows = 0;
        while (reader.Read())
            rows++;
        return rows;
    }

    [Benchmark]
    public int ReadWithComments()
    {
        using var reader = Open();
        int comments = 0;
        while (reader.Read())
        {
            for (int column = 0; column < reader.FieldCount; column++)
                comments += reader.GetComment(column)?.Messages.Count ?? 0;
        }

        return comments;
    }

    [Benchmark]
    public CellCommentThread? OpenAndFirstLookup()
    {
        using var reader = Open();
        reader.Read();
        if (HasComments)
            reader.Read();
        return reader.GetComment(_column);
    }

    [Benchmark(OperationsPerInvoke = 1000)]
    public int CachedLookup()
    {
        int messages = 0;
        for (int i = 0; i < 1000; i++)
            messages += _cachedReader.GetComment(_column)?.Messages.Count ?? 0;
        return messages;
    }

    private IExcelDataReader Open() => ExcelReaderFactory.CreateReader(new MemoryStream(_data, writable: false));
}