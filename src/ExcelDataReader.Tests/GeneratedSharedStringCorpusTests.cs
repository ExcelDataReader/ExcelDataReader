#nullable enable

using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Tests;

public class GeneratedSharedStringCorpusTests
{
    private const long Budget = 1024 * 1024;
    private readonly Dictionary<string, string> _workbookPaths = new();
    private string? _directory;
    private string? _spillDirectory;
    private string[] _expectedValues = Array.Empty<string>();

    [OneTimeSetUp]
    public void CreateGeneratedCorpus()
    {
        _directory = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        _spillDirectory = Path.Combine(_directory, "spill");
        Directory.CreateDirectory(_spillDirectory);
        _expectedValues = Enumerable.Range(0, 20000)
            .Select(position => SharedStringWorkbook.Value(SharedStringWorkbook.Reference(position, 20000, "Permuted"), 64))
            .ToArray();

        foreach (string format in new[] { "xlsx", "xlsb", "xls" })
        {
            string path = Path.Combine(_directory, "corpus." + format);
            SharedStringWorkbook.Create(path, format, 20000, 64, "Permuted");
            _workbookPaths.Add(format, path);
        }
    }

    [OneTimeTearDown]
    public void DeleteGeneratedCorpus()
    {
        if (_directory is not null && Directory.Exists(_directory))
            Directory.Delete(_directory, recursive: true);
    }

    [Test]
    [Combinatorial]
    public void GeneratedCorpus_PreservesValuesAndReset(
        [Values("xlsx", "xlsb", "xls")] string format,
        [Values(SharedStringStorageMode.Default, SharedStringStorageMode.SpillToDisk)] SharedStringStorageMode mode,
        [Values(false, true)] bool singlePass)
    {
        using var input = File.OpenRead(_workbookPaths[format]);
        using (var reader = ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration
        {
            SharedStringStorageMode = mode,
            SharedStringSpillThreshold = Budget,
            SharedStringTemporaryDirectory = _spillDirectory,
            SinglePassMode = singlePass,
            LeaveOpen = true,
        }))
        {
            Assert.That(Directory.GetFiles(_spillDirectory!).Length, Is.EqualTo(mode == SharedStringStorageMode.SpillToDisk ? 2 : 0));
            for (int pass = 0; pass < 2; pass++)
            {
                int position = 0;
                do
                {
                    while (reader.Read())
                    {
                        for (int c = 0; c < reader.FieldCount; c++)
                        {
                            string expected = _expectedValues[position++];
                            string first = reader.GetString(c);
                            string second = reader.GetString(c);
                            if (first != expected || second != expected)
                            {
                                Assert.That(first, Is.EqualTo(expected));
                                Assert.That(second, Is.EqualTo(expected));
                            }
                        }
                    }
                }
                while (reader.NextResult());
                Assert.That(position, Is.EqualTo(20000));
                reader.Reset();
            }
        }

        Assert.That(input.CanRead, Is.True);
        Assert.That(Directory.GetFiles(_spillDirectory!), Is.Empty);
    }
}
