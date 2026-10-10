#nullable enable

using System.IO.Compression;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Xml;
using ExcelDataReader.Exceptions;
using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Tests;

public class SharedStringStorageTests
{
    private const long Budget = 1024 * 1024;

    [Test]
    public void MappedViewsPreserveBoundariesBoundsAndBufferedMissPolicy()
    {
        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        using var file = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose);
        byte[] expected = Enumerable.Range(0, 3 * 65536 + 123).Select(i => (byte)(i * 741)).ToArray();
        file.Write(expected, 0, expected.Length);
        file.Flush();
        using var mapped = new MappedReader(file, expected.Length, 65536);
        var buffer = new byte[32768];
        Assert.That(mapped.Read(65530, buffer, buffer.Length, true), Is.True);
        Assert.That(buffer, Is.EqualTo(expected.Skip(65530).Take(buffer.Length)));
        byte[] before = (byte[])buffer.Clone();
        Assert.That(mapped.Read(0, buffer, buffer.Length, false), Is.False);
        Assert.That(buffer, Is.EqualTo(before));
        foreach (long offset in new long[] { 0, 65530, 131060, expected.Length - buffer.Length })
        {
            Assert.That(mapped.Read(offset, buffer, buffer.Length, true), Is.True);
            Assert.That(buffer, Is.EqualTo(expected.Skip((int)offset).Take(buffer.Length)));
        }

        Assert.Throws<EndOfStreamException>(() => mapped.Read(-1, buffer, 1, true));
        Assert.Throws<EndOfStreamException>(() => mapped.Read(expected.Length, buffer, 1, true));
        Assert.Throws<EndOfStreamException>(() => mapped.Read(long.MaxValue, buffer, 1, true));
        mapped.Dispose();
        mapped.Dispose();
        Assert.Throws<ObjectDisposedException>(() => mapped.Read(0, buffer, 1, true));
        Assert.That(file.CanRead, Is.True);
    }

    [Test]
    public void MappingSetupRejectsTruncatedFiles()
    {
        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        using var file = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose);
        file.WriteByte(1);
        file.Flush();
        Assert.Throws<EndOfStreamException>(() => new MappedReader(file, 2, 65536));
    }

    [TestCase("payload")]
    [TestCase("index")]
    public void TruncationBeforeMappingIsAnErrorAndDeletesOwnedFiles(string kind)
    {
        var files = new Dictionary<string, FileStream>();
        using var store = new Store(SharedStringStorageMode.SpillToDisk, (directory, name) =>
        {
            var file = new FileStream(Path.Combine(directory, Guid.NewGuid().ToString("N")), FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose);
            files.Add(name, file);
            return file;
        });
        store.Add(new string('x', (int)Budget));
        files[kind].SetLength(files[kind].Length - 1);
        Assert.Throws<EndOfStreamException>(store.Seal);
        Assert.That(files.Values.All(file => !File.Exists(file.Name)), Is.True);
        Assert.That(files.Values.All(file => !file.CanRead), Is.True);
        Assert.Throws<ObjectDisposedException>(() => store.Get(0));
    }

    [Test]
    public void MappedPayloadPreservesUnpairedSurrogatesAndNullCodeUnits()
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        const string value = "\uD800\uDC00\uD800\0\uFFFF\uDC00";
        store.Add(value);
        store.Add(new string('x', (int)Budget));
        store.Seal();
        Assert.That(store.Field<object?>("_payloadMapping"), Is.Not.Null);
        Assert.That(store.Get(0), Is.EqualTo(value));
        object mapping = store.Field<object>("_payloadMapping");
        Assert.That(mapping.GetType().GetField("_viewSize", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(mapping), Is.EqualTo((IntPtr.Size == 4 ? 8 : 64) * 1024 * 1024));
    }

    [TestCase("payload")]
    [TestCase("index")]
    public void MappingCreationFailureDisposesAllFiles(string failure)
    {
        var paths = new List<string>();
        var files = new List<FileStream>();
        using var store = new Store(SharedStringStorageMode.SpillToDisk, (directory, kind) =>
        {
            string path = Path.Combine(directory, Guid.NewGuid().ToString("N"));
            paths.Add(path);
            var file = new FileStream(path, FileMode.CreateNew, kind == failure ? FileAccess.Write : FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose);
            files.Add(file);
            return file;
        });
        store.Add(new string('x', (int)Budget));
        Assert.Throws<UnauthorizedAccessException>(store.Seal);
        Assert.That(paths.All(path => !File.Exists(path)), Is.True);
        Assert.That(files.All(file => !file.CanWrite), Is.True);
        Assert.That(store.Field<object?>("_payloadMapping"), Is.Null);
        Assert.That(store.Field<object?>("_indexMapping"), Is.Null);
        Assert.Throws<ObjectDisposedException>(() => store.Get(0));
    }

    [TestCase(-2L, 1)]
    [TestCase(1L, 1)]
    [TestCase(0L, -1)]
    [TestCase(0L, int.MaxValue)]
    public void MappedIndexCorruptionIsRejected(long offset, int length)
    {
        FileStream? index = null;
        using var store = new Store(SharedStringStorageMode.SpillToDisk, (directory, kind) =>
        {
            var file = new FileStream(Path.Combine(directory, Guid.NewGuid().ToString("N")), FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose);
            if (kind == "index")
                index = file;
            return file;
        });
        store.Add(new string('x', (int)Budget));
        index!.Position = 0;
        index.Write(BitConverter.GetBytes(offset), 0, 8);
        index.Write(BitConverter.GetBytes(length), 0, 4);
        store.Seal();
        Assert.That(store.Field<object?>("_indexMapping"), Is.Not.Null);
        Assert.Throws<IOException>(() => store.Get(0));
    }

    [TestCase("payload")]
    [TestCase("index")]
    public void FailedMappedReadsDoNotExposeAnUnreadWindow(string kind)
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        for (int i = 0; i < 10000; i++)
            store.Add(new string((char)('a' + i % 26), 64));
        store.Seal();
        Assert.That(store.Get(0), Is.EqualTo(new string('a', 64)));
        ((IDisposable)store.Field<object>("_" + kind + "Mapping")).Dispose();
        Assert.Throws<ObjectDisposedException>(() => store.Get(10));
        Assert.Throws<ObjectDisposedException>(() => store.Get(20));
        Assert.That(store.Field<int>("_" + kind + "WindowLength"), Is.Zero);
    }

    [Test]
    public void DiskReadWindowsPreserveVariableLengthCodeUnitsAndStayBounded()
    {
        var streams = new List<FaultStream>();
        using var store = new Store(SharedStringStorageMode.SpillToDisk, (_, _) =>
        {
            var stream = new FaultStream();
            streams.Add(stream);
            return stream;
        });
        string[] values = Enumerable.Range(0, 5000)
            .Select(i => i % 257 == 0 ? string.Empty : new string((char)(0xd800 + i % 2048), 100 + i % 501) + i)
            .ToArray();
        foreach (string value in values)
            store.Add(value);
        store.Seal();
        Assert.That(store.Spilled, Is.True);
        long backingBytes = store.ResidentBytes;
        Assert.That(backingBytes, Is.EqualTo(256 + 32768 + 4096 + 64 + 2 * (4096 + 256)));
        Assert.That(store.Field<byte[]>("_buffer").Length, Is.EqualTo(32768));
        Assert.That(store.Field<byte[]>("_entryBuffer").Length, Is.EqualTo(4096));
        foreach (string pattern in new[] { "Sequential", "Permuted", "Random" })
        {
            for (int position = 0; position < values.Length; position++)
            {
                int index = SharedStringWorkbook.Reference(position, values.Length, pattern);
                Assert.That(store.Get(index), Is.EqualTo(values[index]));
                Assert.That(store.ResidentBytes, Is.LessThanOrEqualTo(Budget));
            }
        }

        Assert.That(store.ResidentBytes - store.Field<long>("_cacheBytes"), Is.EqualTo(backingBytes));
    }

    [TestCase("payload", "Truncated")]
    [TestCase("index", "Truncated")]
    [TestCase("payload", "Read")]
    [TestCase("index", "Read")]
    [TestCase("payload", "Seek")]
    [TestCase("index", "Seek")]
    public void FailedReadAheadNeverExposesAnUnreadTail(string kind, string failure)
    {
        var streams = new Dictionary<string, FaultStream>();
        using var store = new Store(SharedStringStorageMode.SpillToDisk, (_, name) =>
        {
            var stream = new FaultStream();
            streams.Add(name, stream);
            return stream;
        });
        for (int i = 0; i < 10000; i++)
            store.Add(new string((char)('a' + i % 26), 64));
        store.Seal();
        Assert.That(store.Get(0), Is.EqualTo(new string('a', 64)));
        if (failure == "Truncated")
            streams[kind].SetLength(kind == "payload" ? 256 : 32);
        streams[kind].FailRead = failure == "Read";
        streams[kind].FailSeek = failure == "Seek";
        if (failure == "Truncated")
        {
            Assert.Throws<EndOfStreamException>(() => store.Get(1));
            Assert.Throws<EndOfStreamException>(() => store.Get(2));
        }
        else
        {
            Assert.Throws<IOException>(() => store.Get(1));
            Assert.Throws<IOException>(() => store.Get(2));
        }

        Assert.That(store.ResidentBytes, Is.LessThanOrEqualTo(Budget));
    }

    [Test]
    public void UnalignedPayloadOffsetsAreRejectedBeforeDecoding()
    {
        var streams = new Dictionary<string, FaultStream>();
        using var store = new Store(SharedStringStorageMode.SpillToDisk, (_, kind) =>
        {
            var stream = new FaultStream();
            streams.Add(kind, stream);
            return stream;
        });
        for (int i = 0; i < 10000; i++)
            store.Add(new string('x', 64));
        store.Seal();
        streams["index"].Position = 0;
        streams["index"].Write(BitConverter.GetBytes(1L), 0, 8);
        Assert.Throws<IOException>(() => store.Get(0));
    }

    [Test]
    public void ReadAheadHandlesShortReadsAndPartialFinalWindows()
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk, (_, _) => new FaultStream { ShortRead = true });
        for (int i = 0; i < 10000; i++)
            store.Add(new string((char)('a' + i % 26), 64));
        store.Seal();
        Assert.That(store.Spilled, Is.True);
        foreach (int index in new[] { 9998, 9999, 0, 1, 9999, 9998 })
            Assert.That(store.Get(index), Is.EqualTo(new string((char)('a' + index % 26), 64)));
    }

    [Test]
    public void ExactStringBudgetDoesNotSpillUntilExceeded()
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        int length = (int)((Budget - 256 - 4 * IntPtr.Size - 32) / 2);
        string value = new('x', length);
        store.Add(value);
        Assert.That(store.ResidentBytes, Is.EqualTo(Budget));
        Assert.That(store.Spilled, Is.False);
        store.Add("next");
        Assert.That(store.Spilled, Is.True);
        Assert.That(store.Field<List<string>?>("_strings"), Is.Null);
        store.Seal();
        Assert.That(store.Get(0), Is.EqualTo(value));
        Assert.That(store.Get(1), Is.EqualTo("next"));
    }

    [TestCase("xlsx")]
    [TestCase("xlsb")]
    [TestCase("xls")]
    public void BelowThresholdAutoMatchesDefaultValuesAndRepeatedReferenceIdentity(string format)
    {
        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") + "." + format);
        string directory = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            SharedStringWorkbook.Create(path, format, 512, 64, "Hot", unicode: true);
            var expected = ReadFixturePath(path, SharedStringStorageMode.Default);
            foreach (var mode in new[] { SharedStringStorageMode.Default, SharedStringStorageMode.SpillToDisk })
            {
                using var input = File.OpenRead(path);
                using var reader = ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration
                {
                    SharedStringStorageMode = mode,
                    SharedStringSpillThreshold = Budget,
                    SharedStringTemporaryDirectory = directory,
                    SinglePassMode = true,
                });
                string? first = null;
                int position = 0;
                while (reader.Read())
                {
                    for (int c = 0; c < reader.FieldCount; c++)
                    {
                        string value = reader.GetString(c);
                        Assert.That(value, Is.EqualTo(expected[position]));
                        if (position == 0)
                            first = value;
                        if (position == 256)
                            Assert.That(value, Is.SameAs(first));
                        position++;
                    }
                }

                Assert.That(position, Is.EqualTo(expected.Count));
                Assert.That(Directory.GetFiles(directory), Is.Empty);
                reader.Reset();
                Assert.That(reader.Read(), Is.True);
                Assert.That(reader.GetString(0), Is.SameAs(first));
            }
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    [TestCase(SharedStringStorageMode.SpillToDisk)]
    public void DecodedCacheAdmitsSecondTouchesAndSurvivesOneOffScan(SharedStringStorageMode mode)
    {
        using var store = new Store(mode);
        for (int i = 0; i < 5000; i++)
            store.Add("value-" + i);
        store.Add(new string('x', (int)Budget));
        store.Seal();
        string first = store.Get(0);
        store.Get(1);
        string admitted = store.Get(0);
        Assert.That(admitted, Is.Not.SameAs(first));
        store.Get(1);
        Assert.That(store.Get(0), Is.SameAs(admitted));
        for (int i = 2; i < 5000; i++)
            store.Get(i);
        Assert.That(store.Get(0), Is.SameAs(admitted));
        int collision = store.Field<int[]>("_cacheIndices").Length;
        store.Get(collision);
        store.Get(2);
        store.Get(collision);
        Assert.That(store.Get(0), Is.EqualTo("value-0").And.Not.SameAs(admitted));
        if (mode == SharedStringStorageMode.SpillToDisk)
            Assert.That(store.ResidentBytes, Is.LessThanOrEqualTo(Budget));
        store.Dispose();
        Assert.That(store.Field<string?[]?>("_cacheValues"), Is.Null);
        Assert.That(store.ResidentBytes, Is.EqualTo(256));
    }

    [Test]
    public void DecodedCacheAccountsMetadataAndValuesExactly()
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        store.Add("first");
        store.Add("\u03bb second");
        store.Add(new string('x', (int)Budget));
        store.Seal();
        long backingBytes = store.ResidentBytes;
        for (int i = 0; i < 6; i++)
            store.Get(i % 2);
        var values = store.Field<string?[]>("_cacheValues");
        long expected = values.Length * 16L + 96;
        foreach (string? value in values)
        {
            if (value != null)
                expected += 32L + value.Length * 2L;
        }

        expected += 32L + store.Field<string>("_cachedValue").Length * 2L;
        Assert.That(store.Field<long>("_cacheBytes"), Is.EqualTo(expected));
        Assert.That(store.ResidentBytes, Is.EqualTo(backingBytes + expected));
    }

    [Test]
    public void DecodedCacheCollisionsEvictionAndLargeValuesRemainBounded()
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        for (int i = 0; i < 5000; i++)
            store.Add(new string('x', i == 4999 ? 40000 : 1000) + i);
        store.Seal();
        Assert.That(store.Spilled, Is.True);
        for (int pass = 0; pass < 4; pass++)
        {
            for (int i = 0; i < 100; i++)
            {
                store.Get(i);
                Assert.That(store.Field<long>("_cacheBytes"), Is.LessThanOrEqualTo(Budget / 8));
            }
        }

        store.Get(4999);
        Assert.That(store.Field<long>("_cacheBytes"), Is.LessThanOrEqualTo(Budget / 8));
        for (int pass = 0; pass < 3; pass++)
        {
            for (int i = 0; i < 5000; i++)
            {
                Assert.That(store.Get(i), Is.EqualTo(new string('x', i == 4999 ? 40000 : 1000) + i));
                Assert.That(store.ResidentBytes, Is.LessThanOrEqualTo(Budget));
                Assert.That(store.Field<long>("_cacheBytes"), Is.LessThanOrEqualTo(Budget / 8));
            }
        }
    }

    [TestCase(1)]
    [TestCase(257)]
    [TestCase(1000)]
    [TestCase(104729)]
    public void RandomCorpusReferencesAreAnExactPermutation(int count)
    {
        var seen = new HashSet<int>();
        for (int i = 0; i < count; i++)
        {
            int index = SharedStringWorkbook.Reference(i, count, "Random");
            Assert.That(index, Is.InRange(0, count - 1));
            Assert.That(seen.Add(index), Is.True);
        }
    }

    [Test]
    public void MalformedBiffSurrogatesRetainDefaultDecoding()
    {
        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") + ".xls");
        string directory = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(directory);
            SharedStringWorkbook.Create(path, "xls", 20000, 64, unicode: true);
            byte[] bytes = File.ReadAllBytes(path);
            for (int offset = 0; offset < bytes.Length; offset += 4 + BitConverter.ToUInt16(bytes, offset + 2))
            {
                if (BitConverter.ToUInt16(bytes, offset) != 0xFC)
                    continue;
                bytes[offset + 15] = 0;
                bytes[offset + 16] = 0xD8;
                break;
            }

            File.WriteAllBytes(path, bytes);
            foreach (SharedStringStorageMode mode in Enum.GetValues(typeof(SharedStringStorageMode)))
            {
                using var input = File.OpenRead(path);
                using var reader = ExcelReaderFactory.CreateBinaryReader(input, new ExcelReaderConfiguration
                {
                    SharedStringStorageMode = mode,
                    SharedStringSpillThreshold = Budget,
                    SharedStringTemporaryDirectory = directory,
                });
                if (mode == SharedStringStorageMode.SpillToDisk)
                    Assert.That(Directory.GetFiles(directory), Has.Length.EqualTo(2));
                Assert.That(reader.Read(), Is.True);
                Assert.That(reader.GetString(0), Is.EqualTo("\uFFFD" + SharedStringWorkbook.Value(0, 64, true).Substring(1)));
            }
        }
        finally
        {
            File.Delete(path);
            if (Directory.Exists(directory))
            {
                foreach (string file in Directory.GetFiles(directory))
                    File.Delete(file);
                Directory.Delete(directory);
            }
        }
    }

    [Test]
    public void GeneratedCorpus_AllFormatsAndModesPreserveValuesAndReset()
    {
        foreach (string format in new[] { "xlsx", "xlsb", "xls" })
        {
            string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") + "." + format);
            string directory = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                SharedStringWorkbook.Create(path, format, 20000, 64, "Permuted");
                foreach (SharedStringStorageMode mode in Enum.GetValues(typeof(SharedStringStorageMode)))
                {
                    foreach (bool singlePass in new[] { false, true })
                    {
                        using var input = File.OpenRead(path);
                        using (var reader = ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration
                        {
                            SharedStringStorageMode = mode,
                            SharedStringSpillThreshold = Budget,
                            SharedStringTemporaryDirectory = directory,
                            SinglePassMode = singlePass,
                            LeaveOpen = true,
                        }))
                        {
                            Assert.That(Directory.GetFiles(directory).Length, Is.EqualTo(mode == SharedStringStorageMode.SpillToDisk ? 2 : 0));
                            for (int pass = 0; pass < 2; pass++)
                            {
                                int position = 0;
                                do
                                {
                                    while (reader.Read())
                                    {
                                        for (int c = 0; c < reader.FieldCount; c++)
                                        {
                                            string expected = SharedStringWorkbook.Value(SharedStringWorkbook.Reference(position++, 20000, "Permuted"), 64);
                                            Assert.That(reader.GetString(c), Is.EqualTo(expected));
                                            Assert.That(reader.GetString(c), Is.EqualTo(expected));
                                        }
                                    }
                                }
                                while (reader.NextResult());
                                Assert.That(position, Is.EqualTo(20000));
                                reader.Reset();
                            }
                        }

                        Assert.That(input.CanRead, Is.True);
                        Assert.That(Directory.GetFiles(directory), Is.Empty);
                    }
                }
            }
            finally
            {
                File.Delete(path);
                Directory.Delete(directory);
            }
        }
    }

    [TestCase(SharedStringStorageMode.SpillToDisk)]
    public void AutoStorePreservesCodeUnitsAndDiskBufferBoundaries(SharedStringStorageMode mode)
    {
        using var store = new Store(mode);
        string[] values = [string.Empty, "\0\u00ff\u0080", "\uD800\uDC00\uD800", new string('a', 40000), new string('\u03BB', 40000)];
        for (int i = 0; i < 20; i++)
        {
            foreach (string value in values)
                store.Add(value);
        }

        store.Seal();
        for (int i = store.Count - 1; i >= 0; i--)
            Assert.That(store.Get(i), Is.EqualTo(values[i % values.Length]));
        Assert.That(store.Spilled, Is.EqualTo(mode == SharedStringStorageMode.SpillToDisk));
        if (mode == SharedStringStorageMode.SpillToDisk)
            Assert.That(store.ResidentBytes, Is.LessThanOrEqualTo(Budget));
        Assert.Throws<ArgumentOutOfRangeException>(() => store.Get(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => store.Get(store.Count));
        Assert.Throws<InvalidOperationException>(() => store.Add("after seal"));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void XlsbLongRecordsPreserveValuesAcrossSpill(bool unicode)
    {
        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") + ".xlsb");
        try
        {
            SharedStringWorkbook.Create(path, "xlsb", 400, 4096, unicode: unicode);
            foreach (SharedStringStorageMode mode in Enum.GetValues(typeof(SharedStringStorageMode)))
            {
                using var input = File.OpenRead(path);
                using var reader = ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration
                {
                    SharedStringStorageMode = mode,
                    SharedStringSpillThreshold = Budget,
                });
                int index = 0;
                while (reader.Read())
                {
                    for (int c = 0; c < reader.FieldCount; c++)
                        Assert.That(reader.GetString(c), Is.EqualTo(SharedStringWorkbook.Value(index++, 4096, unicode)));
                }

                Assert.That(index, Is.EqualTo(400));
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestCase(2, 0xffu, false)]
    [TestCase(5, 0xffu, false)]
    [TestCase(9, 3u, false)]
    [TestCase(5, 0x80000000u, false)]
    [TestCase(5, uint.MaxValue, false)]
    [TestCase(9, 3u, true)]
    public void XlsbRejectsStringLengthsOutsideRecord(int recordLength, uint count, bool afterSpill)
    {
        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") + ".xlsb");
        string directory = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(directory);
            SharedStringWorkbook.Create(path, "xlsb", afterSpill ? 20000 : 1, 64, unicode: afterSpill);
            using (var input = File.Open(path, FileMode.Open, FileAccess.ReadWrite))
            using (var zip = new ZipArchive(input, ZipArchiveMode.Update))
            {
                var entry = zip.GetEntry("xl/sharedStrings.bin")!;
                var data = new MemoryStream();
                if (afterSpill)
                {
                    using var stream = entry.Open();
                    stream.CopyTo(data);
                }

                entry.Delete();
                data.WriteByte(0x13);
                data.WriteByte((byte)recordLength);
                var bytes = new byte[recordLength];
                byte[] countBytes = BitConverter.GetBytes(count);
                Array.Copy(countBytes, 0, bytes, 1, Math.Min(4, recordLength - 1));
                data.Write(bytes, 0, bytes.Length);
                using var output = zip.CreateEntry("xl/sharedStrings.bin").Open();
                data.WriteTo(output);
            }

            foreach (SharedStringStorageMode mode in Enum.GetValues(typeof(SharedStringStorageMode)))
            {
                using var input = File.OpenRead(path);
                var exception = Assert.Throws<ExcelReaderException>(() => ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration
                {
                    SharedStringStorageMode = mode,
                    SharedStringSpillThreshold = Budget,
                    SharedStringTemporaryDirectory = directory,
                }));
                Assert.That(exception!.Message, Is.EqualTo("Error reading BIFF string - record size is too small."));
                Assert.That(Directory.GetFiles(directory), Is.Empty);
            }
        }
        finally
        {
            File.Delete(path);
            if (Directory.Exists(directory))
            {
                foreach (string file in Directory.GetFiles(directory))
                    File.Delete(file);
                Directory.Delete(directory);
            }
        }
    }

    [Test]
    public void XlsbStringThatExactlyFillsRecordIsAccepted()
    {
        var sst = CreateSharedStringTable();
        byte[] text = CodeUnits("ab");
        using var input = new MemoryStream(XlsbStringItem(text, 2));
        var reader = CreateBinarySharedStringsReader(input);
        Assert.Throws<InvalidOperationException>(() => ReadSharedStringsRecord(reader));
        Assert.That(sst, Is.Empty);
        LoadBinarySharedStrings(reader, sst);
        ((IDisposable)reader).Dispose();
        Assert.That(sst, Is.EqualTo(new[] { "ab" }));
    }

    [TestCase("default")]
    [TestCase("resident")]
    [TestCase("spilled")]
    public void XlsbStringItemsAreConsumedBeforeRecordBufferReuse(string target)
    {
        string[] values =
        [
            new string('a', 100),
            new string('b', 100),
            SharedStringWorkbook.Value(1, 1000, true),
            SharedStringWorkbook.Value(2, 1000, true),
            "rich\u03bb",
            string.Empty,
            "\ud83d\ude00\u0000",
            "\ud800x\udc00",
            new string('a', 100),
        ];
        var input = new MemoryStream();
        input.Write([0x9f, 0x01, 8, 0, 0, 0, 0, 0, 0, 0, 0], 0, 11);
        foreach (string value in values)
        {
            byte[] text = CodeUnits(value);
            bool rich = value.StartsWith("rich", StringComparison.Ordinal);
            byte[] record = XlsbStringItem(text, value.Length, rich ? (byte)1 : (byte)0, rich ? 14 : 0);
            input.Write(record, 0, record.Length);
        }

        input.Position = 0;
        List<string>? sst = target == "default" ? CreateSharedStringTable() : null;
        using var store = target == "default" ? null : new Store(SharedStringStorageMode.SpillToDisk);
        if (target == "spilled")
            store!.Add(new string('x', (int)Budget));
        int first = store?.Count ?? 0;
        var sink = (object?)sst ?? store!.Instance;
        var reader = CreateBinarySharedStringsReader(input);
        LoadBinarySharedStrings(reader, sink);
        ((IDisposable)reader).Dispose();
        string[] expected = values.Select(value => Encoding.Unicode.GetString(CodeUnits(value))).ToArray();
        Assert.That(expected[7], Is.EqualTo("\ufffdx\ufffd"));
        if (sst != null)
        {
            Assert.That(sst, Is.EqualTo(expected));
            return;
        }

        Assert.That(store!.Spilled, Is.EqualTo(target == "spilled"));
        Assert.That(store.Count, Is.EqualTo(first + values.Length));
        store.Seal();
        for (int i = 0; i < expected.Length; i++)
            Assert.That(store.Get(first + i), Is.EqualTo(expected[i]));
    }

    [TestCase("resident")]
    [TestCase("crossing")]
    [TestCase("spilled")]
    public void Utf16IngestionCopiesBorrowedBuffersAndMatchesStringIngestion(string state)
    {
        var utf16Streams = new List<MemoryStream>();
        var stringStreams = new List<MemoryStream>();
        Func<string, string, Stream> Capture(List<MemoryStream> streams) => (_, _) =>
        {
            var stream = new MemoryStream();
            streams.Add(stream);
            return stream;
        };
        using var utf16 = new Store(SharedStringStorageMode.SpillToDisk, Capture(utf16Streams));
        using var strings = new Store(SharedStringStorageMode.SpillToDisk, Capture(stringStreams));
        string prefix = state switch
        {
            "spilled" => new string('x', (int)Budget),
            "crossing" => new string('x', (int)((Budget - 330) / 2)),
            _ => "prefix",
        };
        utf16.Add(prefix);
        strings.Add(prefix);
        string[] values = [string.Empty, "BIFF12\u00ff\u03bb", "\ud83d\ude00\u0000", "\ud800", "x\udc00", "\ud800\ud800\udc00", SharedStringWorkbook.Value(3, 1500, true)];
        byte[] buffer = new byte[4096];
        foreach (string value in values)
        {
            byte[] bytes = CodeUnits(value);
            Array.Copy(bytes, 0, buffer, 7, bytes.Length);
            utf16.Invoke("AddUtf16", buffer, 7, value.Length);
            for (int i = 0; i < buffer.Length; i++)
                buffer[i] = 0xff;
            strings.Add(Encoding.Unicode.GetString(bytes));
            Assert.That(utf16.Count, Is.EqualTo(strings.Count));
            Assert.That(utf16.ResidentBytes, Is.EqualTo(strings.ResidentBytes));
            Assert.That(utf16.Spilled, Is.EqualTo(strings.Spilled));
        }

        Assert.That(utf16.Spilled, Is.EqualTo(state != "resident"));
        foreach (FieldInfo field in utf16.Instance.GetType().GetFields(BindingFlags.Instance | BindingFlags.NonPublic))
            Assert.That(field.GetValue(utf16.Instance), Is.Not.SameAs(buffer), field.Name);
        utf16.Seal();
        strings.Seal();
        Assert.That(utf16.DiskBytes, Is.EqualTo(strings.DiskBytes));
        Assert.That(utf16.ResidentBytes, Is.EqualTo(strings.ResidentBytes));
        Assert.That(utf16Streams.Select(stream => stream.ToArray()), Is.EqualTo(stringStreams.Select(stream => stream.ToArray())));
        for (int i = 0; i < values.Length; i++)
            Assert.That(utf16.Get(i + 1), Is.EqualTo(Encoding.Unicode.GetString(CodeUnits(values[i]))));
        if (state == "resident")
        {
            string identity = utf16.Get(2);
            Assert.That(utf16.Get(2), Is.SameAs(identity));
        }
    }

    [TestCase(-1, 0)]
    [TestCase(0, -1)]
    [TestCase(1, 4)]
    [TestCase(9, 1)]
    public void Utf16IngestionRejectsRangesOutsideBuffer(int offset, int count)
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        Assert.Throws<ArgumentOutOfRangeException>(() => store.Invoke("AddUtf16", new byte[8], offset, count));
        Assert.That(store.Count, Is.Zero);
    }

    [Test]
    public void MalformedXlsbSurrogatesRetainDefaultDecoding()
    {
        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") + ".xlsb");
        string directory = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(directory);
            SharedStringWorkbook.Create(path, "xlsb", 20000, 64, unicode: true);
            string expected;
            using (var input = File.Open(path, FileMode.Open, FileAccess.ReadWrite))
            using (var zip = new ZipArchive(input, ZipArchiveMode.Update))
            {
                var entry = zip.GetEntry("xl/sharedStrings.bin")!;
                byte[] bytes;
                using (var data = new MemoryStream())
                {
                    using (var stream = entry.Open())
                        stream.CopyTo(data);
                    bytes = data.ToArray();
                }

                Assert.That(bytes[11], Is.EqualTo(0x13));
                bytes[19] = 0;
                bytes[20] = 0xD8;
                bytes[21] = 0;
                bytes[22] = 0xD8;
                bytes[23] = 0;
                bytes[24] = 0xDC;
                bytes[25] = 0;
                bytes[26] = 0xDC;
                expected = Encoding.Unicode.GetString(bytes, 19, 128);
                entry.Delete();
                using var output = zip.CreateEntry("xl/sharedStrings.bin").Open();
                output.Write(bytes, 0, bytes.Length);
            }

            foreach (SharedStringStorageMode mode in Enum.GetValues(typeof(SharedStringStorageMode)))
            {
                using var input = File.OpenRead(path);
                using var reader = ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration
                {
                    SharedStringStorageMode = mode,
                    SharedStringSpillThreshold = Budget,
                    SharedStringTemporaryDirectory = directory,
                });
                if (mode == SharedStringStorageMode.SpillToDisk)
                    Assert.That(Directory.GetFiles(directory), Has.Length.EqualTo(2));
                Assert.That(reader.Read(), Is.True);
                Assert.That(reader.GetString(0), Is.EqualTo(expected));
                Assert.That(reader.GetString(1), Is.EqualTo(SharedStringWorkbook.Value(1, 64, true)));
            }
        }
        finally
        {
            File.Delete(path);
            if (Directory.Exists(directory))
            {
                foreach (string file in Directory.GetFiles(directory))
                    File.Delete(file);
                Directory.Delete(directory);
            }
        }
    }

    [TestCase(false)]
    [TestCase(true)]
    public void XmlSharedStringParsingMatchesDefault(bool strict)
    {
        string boundary = new string('a', 4095) + "\uD83D\uDE00" + new string('b', 40000);
        string[] items =
        [
            string.Empty,
            "<t/>",
            "<t> \t\r\n </t>",
            "<t> \u00a0value\u00a0 </t>",
            "<t>  first<!--ignored--><![CDATA[ & second ]]><?test ignored?> third  </t>",
            "<r><rPr><b/></rPr><t> first </t></r><rPh sb=\"0\" eb=\"1\"><t>ignored</t></rPh><r><t xml:space=\"preserve\"> second </t></r><t> third </t>",
            "<t>_xD800_ _x005F_x0041_ &amp; &lt; &#x1F600;</t>",
            "<t>" + boundary + "</t>",
            "<r><t>" + new string(' ', 5000) + "left" + new string(' ', 5000) + "</t></r><r><t>" + new string('\t', 5000) + "right" + new string('\t', 5000) + "</t></r>",
        ];
        var all = items.Concat(Enumerable.Repeat("<t>" + boundary + "</t>", 30)).Concat(items).ToArray();
        byte[] data = AllocationTestWorkbook.CreateXlsx(all, true);
        if (strict)
        {
            using var input = new MemoryStream();
            input.Write(data, 0, data.Length);
            input.Position = 0;
            using (var zip = new ZipArchive(input, ZipArchiveMode.Update, leaveOpen: true))
            {
                var entry = zip.GetEntry("xl/sharedStrings.xml")!;
                string xml;
                using (var reader = new StreamReader(entry.Open()))
                    xml = reader.ReadToEnd();
                entry.Delete();
                using var writer = new StreamWriter(zip.CreateEntry("xl/sharedStrings.xml").Open());
                writer.Write(xml.Replace("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main"));
            }

            data = input.ToArray();
        }

        using var baseline = ExcelReaderFactory.CreateReader(new MemoryStream(data));
        var expected = new List<string>();
        while (baseline.Read())
            expected.Add(baseline.GetString(0));
        Assert.That(expected[3], Is.EqualTo("\u00a0value\u00a0"));
        Assert.That(expected[4], Is.EqualTo("first & second  third"));
        Assert.That(expected[5], Is.EqualTo("first second third"));
        Assert.That(expected[7], Is.EqualTo(boundary));
        Assert.That(expected[8], Is.EqualTo("leftright"));
        foreach (var mode in new[] { SharedStringStorageMode.SpillToDisk })
        {
            using var reader = ExcelReaderFactory.CreateReader(
                new MemoryStream(data),
                new ExcelReaderConfiguration { SharedStringStorageMode = mode, SharedStringSpillThreshold = Budget });
            for (int pass = 0; pass < 2; pass++)
            {
                int index = 0;
                while (reader.Read())
                    Assert.That(reader.GetString(0), Is.EqualTo(expected[index++]));
                Assert.That(index, Is.EqualTo(expected.Count));
                reader.Reset();
            }
        }
    }

    [TestCase(SharedStringStorageMode.Default)]
    [TestCase(SharedStringStorageMode.SpillToDisk)]
    public void NestedXmlTextContentRemainsAnError(SharedStringStorageMode mode)
    {
        Assert.Throws<XmlException>(() => ExcelReaderFactory.CreateReader(
            new MemoryStream(AllocationTestWorkbook.CreateXlsx(["<t>text<b/>nested</t>"], true)),
            new ExcelReaderConfiguration { SharedStringStorageMode = mode }));
    }

    [Test]
    public void NormalTablePreservesMixedCodeUnitsAndIdentity()
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        string[] values = [string.Empty, "\0\u00ff", "\u03bb", "\uD800", "plain"];
        for (int i = 0; i < 6145; i++)
            store.Add(values[i % values.Length]);
        store.Seal();
        for (int i = store.Count - 1; i >= 0; i--)
            Assert.That(store.Get(i), Is.SameAs(values[i % values.Length]));
        Assert.That(store.Spilled, Is.False);
        Assert.That(store.Field<string?[]?>("_cacheValues"), Is.Null);
        Assert.That(store.Field<byte[]?>("_buffer"), Is.Null);
    }

    [TestCase(1, 4)]
    [TestCase(4, 4)]
    [TestCase(5, 8)]
    [TestCase(8192, 8192)]
    [TestCase(8193, 16384)]
    public void NormalTableAccountsAllocatedListCapacity(int count, int capacity)
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        for (int i = 0; i < count; i++)
            store.Add(string.Empty);
        Assert.That(store.ResidentBytes, Is.EqualTo(256L + capacity * (long)IntPtr.Size));
        store.Dispose();
        Assert.That(store.Field<List<string>?>("_strings"), Is.Null);
    }

    [Test]
    public void NormalTableSpillsBeforeReferenceCapacityExceedsBudget()
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        int capacity = (int)(Budget / IntPtr.Size / 2);
        for (int i = 0; i < capacity; i++)
            store.Add(string.Empty);
        Assert.That(store.Spilled, Is.False);
        Assert.That(store.ResidentBytes, Is.EqualTo(256L + capacity * (long)IntPtr.Size));
        store.Add(string.Empty);
        Assert.That(store.Spilled, Is.True);
        Assert.That(store.ResidentBytes, Is.LessThan(Budget));
        Assert.That(store.Field<List<string>?>("_strings"), Is.Null);
    }

    [Test]
    public void DiskStorageUsesUniformUtf16AndSixteenByteIndexRecords()
    {
        var streams = new List<MemoryStream>();
        using var store = new Store(SharedStringStorageMode.SpillToDisk, (directory, kind) =>
        {
            var stream = new MemoryStream();
            streams.Add(stream);
            return stream;
        });
        store.Add(string.Empty);
        store.Add("\u00ff");
        store.Add("\u03bb");
        store.Add(new string('x', (int)Budget * 2));
        store.Seal();
        byte[] index = streams[1].ToArray();
        Assert.That(index.Length, Is.EqualTo(4 * 16));
        long[] offsets = [0, 0, 2, 4];
        int[] lengths = [0, 1, 1, (int)Budget * 2];
        for (int i = 0; i < offsets.Length; i++)
        {
            Assert.That(BitConverter.ToInt64(index, i * 16), Is.EqualTo(offsets[i]));
            Assert.That(BitConverter.ToInt32(index, i * 16 + 8), Is.EqualTo(lengths[i]));
            Assert.That(index.Skip(i * 16 + 12).Take(4), Is.All.Zero);
        }

        Assert.That(streams[0].ToArray().Take(4), Is.EqualTo(new byte[] { 0xff, 0, 0xbb, 0x03 }));
        Assert.That(store.DiskBytes, Is.EqualTo(4L + Budget * 4 + 4 * 16));
        Assert.That(store.Get(2), Is.EqualTo("\u03bb"));
        Assert.That(store.Get(1), Is.EqualTo("\u00ff"));
    }

    [Test]
    public void DiskIndexRetainsFullLengthAnd64BitOffsets()
    {
        var streams = new List<MemoryStream>();
        using var store = new Store(SharedStringStorageMode.SpillToDisk, (_, _) =>
        {
            var stream = new MemoryStream();
            streams.Add(stream);
            return stream;
        });
        store.Add(new string('x', (int)Budget));
        Type storeType = typeof(ExcelReaderFactory).Assembly.GetType("ExcelDataReader.Core.AutoSharedStringStore")!;
        Type entryType = storeType.GetNestedType("Entry", BindingFlags.NonPublic)!;
        long offset = (1L << 32) + 741;
        object entry = entryType.GetConstructors()[0].Invoke([offset, int.MaxValue]);
        store.Invoke("WriteEntry", entry);
        byte[] bytes = streams[1].ToArray();
        Assert.That(bytes.Length, Is.EqualTo(32));
        Assert.That(BitConverter.ToInt64(bytes, 16), Is.EqualTo(offset));
        Assert.That(BitConverter.ToInt32(bytes, 24), Is.EqualTo(int.MaxValue));
        Assert.That(bytes.Skip(28).Take(4), Is.All.Zero);
    }

    [Test]
    public void IndexOnlyGrowthSpillsAndStaysBounded()
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        for (int i = 0; i < 200000; i++)
        {
            store.Add(string.Empty);
            Assert.That(store.ResidentBytes, Is.LessThanOrEqualTo(Budget));
        }

        store.Seal();
        Assert.That(store.Spilled, Is.True);
        Assert.That(store.DiskBytes, Is.EqualTo(200000L * 16));
        foreach (int index in new[] { 0, 2047, 2048, 100000, 199999 })
            Assert.That(store.Get(index), Is.Empty);
    }

    [Test]
    public void ThresholdChecksAllocatedCapacityBeforeGrowth()
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        while (!store.Spilled)
        {
            Assert.That(store.ResidentBytes, Is.LessThanOrEqualTo(Budget));
            store.Add(new string('x', 4096));
        }

        Assert.That(store.ResidentBytes, Is.LessThanOrEqualTo(Budget));
        for (int i = 0; i < 1000; i++)
            store.Add("last");
        store.Seal();
        Assert.That(store.Get(0), Has.Length.EqualTo(4096));
        Assert.That(store.Get(store.Count - 1), Is.EqualTo("last"));
        Assert.That(store.ResidentBytes, Is.LessThanOrEqualTo(Budget));
    }

    [Test]
    public void RandomDiskLookupsAndCacheRemainBounded()
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        for (int i = 0; i < 20000; i++)
            store.Add(SharedStringWorkbook.Value(i, 64));
        store.Seal();
        Assert.That(store.Spilled, Is.True);
        var random = new Random(741);
        for (int i = 0; i < 5000; i++)
        {
            int index = random.Next(store.Count);
            string value = store.Get(index);
            Assert.That(value, Is.EqualTo(SharedStringWorkbook.Value(index, 64)));
            Assert.That(store.Get(index), Is.SameAs(value));
            Assert.That(store.ResidentBytes, Is.LessThanOrEqualTo(Budget));
        }
    }

    [TestCase(SharedStringStorageMode.SpillToDisk)]
    public void DeclaredCountDoesNotAllocateHugeIndex(SharedStringStorageMode mode)
    {
        using var input = new MemoryStream();
        byte[] data = AllocationTestWorkbook.CreateXlsx(["<t>value</t>"], true);
        input.Write(data, 0, data.Length);
        input.Position = 0;
        using (var zip = new ZipArchive(input, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("xl/sharedStrings.xml")!;
            string contents;
            using (var reader = new StreamReader(entry.Open()))
                contents = reader.ReadToEnd();
            entry.Delete();
            using var writer = new StreamWriter(zip.CreateEntry("xl/sharedStrings.xml").Open());
            writer.Write(contents.Replace("<sst ", "<sst uniqueCount=\"2147483647\" "));
        }

        input.Position = 0;
        using var excel = ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration
        {
            SharedStringStorageMode = mode,
            SharedStringSpillThreshold = Budget,
        });
        Assert.That(excel.Read(), Is.True);
        Assert.That(excel.GetString(0), Is.EqualTo("value"));
    }

    [Test]
    public void OversizedValueBypassesCacheAndSpills()
    {
        using var store = new Store(SharedStringStorageMode.SpillToDisk);
        store.Add(new string('x', (int)Budget * 2));
        store.Seal();
        Assert.That(store.Spilled, Is.True);
        Assert.That(store.Get(0), Has.Length.EqualTo((int)Budget * 2));
        Assert.That(store.ResidentBytes, Is.LessThanOrEqualTo(Budget));
    }

    [TestCase(SharedStringStorageMode.SpillToDisk)]
    public void SharedXmlParsingAndEscapesAreUnchanged(SharedStringStorageMode mode)
    {
        string[] items = ["<t/>", "<t xml:space=\"preserve\"> \u00ff </t>", "<t>_xD800_</t>", "<r><t>first</t></r><r><t>second</t></r>"];
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(
            new MemoryStream(AllocationTestWorkbook.CreateXlsx(items, true)),
            new ExcelReaderConfiguration { SharedStringStorageMode = mode });
        foreach (string expected in new[] { string.Empty, " \u00ff ", "\uD800", "firstsecond" })
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo(expected));
        }
    }

    [TestCase("Issue286_SST.xls")]
    [TestCase("Issue467_SstEmptyContinue.xls")]
    [TestCase("Issue467_EmptyContinueLeftoverbytes.xls")]
    [TestCase("Issue477_SstWrongCount.xls")]
    [TestCase("Issue477_SstZeroCount.xls")]
    [TestCase("UnicodeChars.xls")]
    [TestCase("MultiSheet.xlsx")]
    [TestCase("MultiSheet.xlsb")]
    [TestCase("MultiSheet.xls")]
    [TestCase("agile_AES128_SHA1_CBC_pwd_password.xlsx")]
    [TestCase("agile_AES128_SHA1_CBC_pwd_password.xlsb")]
    [TestCase("Issue242_StdRc4PwdPassword.xls")]
    public void ExistingFixturesMatchDefaultAcrossModes(string file)
    {
        var expected = ReadFixture(file, SharedStringStorageMode.Default);
        foreach (var mode in new[] { SharedStringStorageMode.SpillToDisk })
            Assert.That(ReadFixture(file, mode), Is.EqualTo(expected));
    }

    [Test]
    public void InvalidBudgetAndDiskCreationFailuresPropagate()
    {
        using var input = new MemoryStream(AllocationTestWorkbook.CreateXlsx(["<t>value</t>"], true));
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration
        {
            SharedStringStorageMode = SharedStringStorageMode.SpillToDisk,
            SharedStringSpillThreshold = Budget - 1,
            LeaveOpen = true,
        }));
        Assert.That(input.CanRead, Is.True);
        input.Position = 0;
        using var reader = ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration
        {
            SharedStringStorageMode = SharedStringStorageMode.SpillToDisk,
            SharedStringTemporaryDirectory = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N")),
            LeaveOpen = true,
        });
        Assert.That(reader.Read(), Is.True);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void FormatSpecificFactorySpillHonorsLeaveOpen(bool leaveOpen)
    {
        foreach (string format in new[] { "xlsx", "xlsb", "xls" })
        {
            string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") + "." + format);
            try
            {
                SharedStringWorkbook.Create(path, format, 20000, 64);
                using var input = File.OpenRead(path);
                var options = new ExcelReaderConfiguration
                {
                    SharedStringStorageMode = SharedStringStorageMode.SpillToDisk,
                    SharedStringSpillThreshold = Budget,
                    LeaveOpen = leaveOpen,
                };
                using var reader = format == "xls"
                    ? ExcelReaderFactory.CreateBinaryReader(input, options)
                    : ExcelReaderFactory.CreateOpenXmlReader(input, options);
                var data = reader.AsDataSet();
                Assert.That(data.Tables[0].Rows.Count, Is.EqualTo(5000));
                Assert.That(data.Tables[0].Rows[0][0], Is.EqualTo(SharedStringWorkbook.Value(0, 64)));
                reader.Close();
                reader.Close();
                Assert.That(input.CanRead, Is.EqualTo(leaveOpen));
            }
            finally
            {
                File.Delete(path);
            }
        }
    }

    [TestCase("Styles")]
    [TestCase("Sheet")]
    public void ConstructorFailureAfterSpillDeletesFiles(string phase)
    {
        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") + ".xlsx");
        string directory = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            SharedStringWorkbook.Create(path, "xlsx", 20000, 64);
            using (var file = File.Open(path, FileMode.Open, FileAccess.ReadWrite))
            using (var zip = new ZipArchive(file, ZipArchiveMode.Update))
            {
                if (phase == "Styles")
                {
                    var rels = zip.GetEntry("xl/_rels/workbook.xml.rels")!;
                    string contents;
                    using (var reader = new StreamReader(rels.Open()))
                        contents = reader.ReadToEnd();
                    rels.Delete();
                    using (var writer = new StreamWriter(zip.CreateEntry("xl/_rels/workbook.xml.rels").Open()))
                        writer.Write(contents.Replace("</Relationships>", "<Relationship Id=\"styles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>"));
                }
                else
                {
                    zip.GetEntry("xl/worksheets/sheet1.xml")!.Delete();
                }

                string entry = phase == "Styles" ? "xl/styles.xml" : "xl/worksheets/sheet1.xml";
                using var malformed = new StreamWriter(zip.CreateEntry(entry).Open(), Encoding.UTF8);
                malformed.Write(phase == "Styles"
                    ? "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><"
                    : "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><");
            }

            using var input = File.OpenRead(path);
            Assert.Throws<System.Xml.XmlException>(() => ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration
            {
                SharedStringStorageMode = SharedStringStorageMode.SpillToDisk,
                SharedStringSpillThreshold = Budget,
                SharedStringTemporaryDirectory = directory,
                LeaveOpen = true,
            }));
            Assert.That(Directory.GetFiles(directory), Is.Empty);
            Assert.That(input.CanRead, Is.True);
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    [Test]
    public void SpillToMissingDirectoryIsNotAnInMemoryFallback()
    {
        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            SharedStringWorkbook.Create(path, "xlsx", 20000, 64);
            using var input = File.OpenRead(path);
            Assert.Throws<DirectoryNotFoundException>(() => ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration
            {
                SharedStringStorageMode = SharedStringStorageMode.SpillToDisk,
                SharedStringSpillThreshold = Budget,
                SharedStringTemporaryDirectory = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N")),
                LeaveOpen = true,
            }));
            Assert.That(input.CanRead, Is.True);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task ConcurrentReadersOwnIndependentSpillFiles()
    {
        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") + ".xlsx");
        string directory = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            SharedStringWorkbook.Create(path, "xlsx", 20000, 64);
            var options = new ExcelReaderConfiguration
            {
                SharedStringStorageMode = SharedStringStorageMode.SpillToDisk,
                SharedStringSpillThreshold = Budget,
                SharedStringTemporaryDirectory = directory,
                SinglePassMode = true,
            };
            await Task.WhenAll(Enumerable.Range(0, 2).Select(_ => Task.Run(() =>
            {
                using var input = File.OpenRead(path);
                using var reader = ExcelReaderFactory.CreateReader(input, options);
                int position = 0;
                while (reader.Read())
                {
                    for (int c = 0; c < reader.FieldCount; c++)
                        Assert.That(reader.GetString(c), Is.EqualTo(SharedStringWorkbook.Value(position++, 64)));
                }

                Assert.That(position, Is.EqualTo(20000));
            })));
            Assert.That(Directory.GetFiles(directory), Is.Empty);
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    [TestCase("Create")]
    [TestCase("Write")]
    [TestCase("Flush")]
    [TestCase("Read")]
    [TestCase("ShortRead")]
    [TestCase("Truncated")]
    [TestCase("Seek")]
    [TestCase("Dispose")]
    [TestCase("Corrupt")]
    public void DiskFailuresRemainErrorsAndDisposeOwnedStreams(string failure)
    {
        var streams = new List<FaultStream>();
        using var store = new Store(SharedStringStorageMode.SpillToDisk, (directory, kind) =>
        {
            if (failure == "Create" && kind == "index")
                throw new IOException("injected create");
            var stream = new FaultStream();
            streams.Add(stream);
            return stream;
        });
        if (failure == "Write")
        {
            store.Add("first");

            // Fail migration of the already populated payload, not a later lookup.
            store.FactoryAction = stream => stream.FailWrite = true;
        }

        if (failure is "Create" or "Write")
        {
            Assert.Throws<IOException>(() => store.Add(new string('x', (int)Budget * 2)));
        }
        else
        {
            store.Add(new string('x', (int)Budget * 2));
            foreach (var stream in streams)
                stream.FailFlush = failure == "Flush";
            if (failure == "Flush")
            {
                Assert.Throws<IOException>(store.Seal);
            }
            else
            {
                store.Seal();
                foreach (var stream in streams)
                {
                    stream.FailRead = failure == "Read";
                    stream.ShortRead = failure == "ShortRead";
                    stream.FailSeek = failure == "Seek";
                }

                if (failure == "Truncated")
                    streams[1].SetLength(0);
                if (failure == "Corrupt")
                {
                    streams[1].Position = 0;
                    byte[] invalidOffset = Enumerable.Repeat((byte)0xff, 8).ToArray();
                    streams[1].Write(invalidOffset, 0, invalidOffset.Length);
                }

                if (failure == "Read")
                {
                    Assert.Throws<IOException>(() => store.Get(0));
                }
                else if (failure == "Truncated")
                {
                    Assert.Throws<EndOfStreamException>(() => store.Get(0));
                }
                else if (failure == "Seek")
                {
                    Assert.Throws<IOException>(() => store.Get(0));
                }
                else if (failure == "Corrupt")
                {
                    Assert.Throws<IOException>(() => store.Get(0));
                }
                else
                {
                    Assert.That(store.Get(0), Has.Length.EqualTo((int)Budget * 2));
                }
            }
        }

        if (failure == "Dispose")
        {
            streams[0].FailDispose = true;
            Assert.Throws<IOException>(store.Dispose);
        }
        else
        {
            store.Dispose();
        }

        Assert.That(streams.All(s => s.WasDisposed), Is.True);
        store.Dispose();
        Assert.Throws<ObjectDisposedException>(() => store.Get(0));
    }

    [Test]
    public void MigrationFailureAndCleanupFailureBothSurface()
    {
        var streams = new List<FaultStream>();
        using var store = new Store(SharedStringStorageMode.SpillToDisk, (directory, kind) =>
        {
            var stream = new FaultStream { FailWrite = kind == "payload", FailDispose = kind == "payload" };
            streams.Add(stream);
            return stream;
        });
        store.Add("first");
        var error = Assert.Throws<AggregateException>(() => store.Add(new string('x', (int)Budget * 2)))!;
        Assert.That(error.Flatten().InnerExceptions.Select(e => e.Message), Does.Contain("injected write").And.Contain("injected dispose"));
        Assert.That(streams.All(s => s.WasDisposed), Is.True);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void XlsSpillConsumesReusedAssemblyBuffer(bool unicode)
    {
        const int entries = 10000;
        const int length = 128;
        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") + ".xls");
        string directory = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            SharedStringWorkbook.Create(path, "xls", entries, length, unicode: unicode);
            using (var input = File.OpenRead(path))
            using (var reader = ExcelReaderFactory.CreateBinaryReader(input, new ExcelReaderConfiguration
            {
                SharedStringStorageMode = SharedStringStorageMode.SpillToDisk,
                SharedStringSpillThreshold = Budget,
                SharedStringTemporaryDirectory = directory,
                SinglePassMode = true,
            }))
            {
                object workbook = reader.GetType().GetProperty("Workbook", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(reader)!;
                object sst = workbook.GetType().GetProperty("SST")!.GetValue(workbook)!;
                object parser = sst.GetType().GetField("_reader", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(sst)!;
                byte[] assemblyBuffer = (byte[])parser.GetType().GetProperty("CurrentResult", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(parser)!;
                Assert.That(assemblyBuffer, Is.Empty);
                int index = 0;
                while (reader.Read())
                {
                    for (int column = 0; column < reader.FieldCount; column++)
                    {
                        Assert.That(reader.GetString(column), Is.EqualTo(SharedStringWorkbook.Value(index, length, unicode)), $"Entry {index}");
                        index++;
                    }
                }

                Assert.That(index, Is.EqualTo(entries));
                Assert.That(Directory.GetFiles(directory), Has.Length.EqualTo(2));
            }

            Assert.That(Directory.GetFiles(directory), Is.Empty);
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    [TestCase(false)]
    [TestCase(true)]
    public void XlsDefaultEagerlyStoresDecodedStrings(bool unicode)
    {
        const int entries = 8;
        const int length = 32;
        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N") + ".xls");
        try
        {
            SharedStringWorkbook.Create(path, "xls", entries, length, unicode: unicode);
            using var input = File.OpenRead(path);
            using var reader = ExcelReaderFactory.CreateBinaryReader(input, new ExcelReaderConfiguration { SinglePassMode = true });
            object workbook = reader.GetType().GetProperty("Workbook", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(reader)!;
            object sst = workbook.GetType().GetProperty("SST")!.GetValue(workbook)!;
            object parser = sst.GetType().GetField("_reader", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(sst)!;
            byte[] assemblyBuffer = (byte[])parser.GetType().GetProperty("CurrentResult", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(parser)!;
            Assert.That(assemblyBuffer, Is.Empty);
            object store = workbook.GetType().GetField("_stringStore", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(workbook)!;
            var getString = store.GetType().GetMethod("GetString")!;
            for (int i = 0; i < entries; i++)
                Assert.That(getString.Invoke(store, [i]), Is.EqualTo(SharedStringWorkbook.Value(i, length, unicode)), $"Entry {i}");

            Assert.That(reader.Read(), Is.True);
            string value = reader.GetString(0);
            Assert.That(value, Is.EqualTo(SharedStringWorkbook.Value(0, length, unicode)));
            Assert.That(getString.Invoke(store, [0]), Is.SameAs(value));
            Assert.That(reader.GetString(0), Is.SameAs(value));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public void XlsDefaultConsumesReusedBufferUsingEachStringsLogicalLength()
    {
        (string Value, bool Compressed)[] values =
        [
            ("long value", false),
            ("x", true),
            (string.Empty, true),
            ("n\u00ff", true),
            ("\ud800", false),
            ("\ud83d\ude00", false),
            ("\0", false),
        ];
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.Unicode, leaveOpen: true))
        {
            writer.Write((ushort)0x00fc);
            writer.Write((ushort)0);
            writer.Write((uint)values.Length);
            writer.Write((uint)values.Length);
            foreach ((string value, bool compressed) in values)
            {
                writer.Write((ushort)value.Length);
                writer.Write((byte)(compressed ? 0 : 1));
                foreach (char codeUnit in value)
                {
                    if (compressed)
                        writer.Write((byte)codeUnit);
                    else
                        writer.Write((ushort)codeUnit);
                }
            }

            writer.Flush();
            stream.Position = 2;
            writer.Write((ushort)(stream.Length - 4));
        }

        byte[] record = stream.ToArray();
        Type sstType = typeof(ExcelReaderFactory).Assembly.GetType("ExcelDataReader.Core.BinaryFormat.XlsBiffSST")!;
        Type tableType = typeof(ExcelReaderFactory).Assembly.GetType("ExcelDataReader.Core.SharedStringTable")!;
        object table = Activator.CreateInstance(tableType)!;
        object sst = sstType.GetConstructors(BindingFlags.Instance | BindingFlags.NonPublic).Single().Invoke([record, table]);
        sstType.GetMethod("Flush")!.Invoke(sst, null);
        var getString = tableType.GetMethod("GetString")!;
        string[] expected = ["long value", "x", string.Empty, "n\u00ff", "\ufffd", "\ud83d\ude00", "\0"];
        for (int i = 0; i < expected.Length; i++)
            Assert.That(getString.Invoke(table, [i]), Is.EqualTo(expected[i]), $"Entry {i}");

        string first = (string)getString.Invoke(table, [0])!;
        Assert.That(getString.Invoke(table, [0]), Is.SameAs(first));
        object parser = sstType.GetField("_reader", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(sst)!;
        byte[] assemblyBuffer = (byte[])parser.GetType().GetProperty("CurrentResult", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(parser)!;
        Assert.That(assemblyBuffer, Is.Empty);
    }

    private static List<string> ReadFixture(string file, SharedStringStorageMode mode)
    {
        using var input = Configuration.GetTestWorkbook(file);
        using var reader = ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration
        {
            Password = "password",
            SharedStringStorageMode = mode,
            SharedStringSpillThreshold = Budget,
        });
        var values = new List<string>();
        do
        {
            while (reader.Read())
            {
                for (int c = 0; c < reader.FieldCount; c++)
                    values.Add(Convert.ToString(reader.GetValue(c), System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
            }
        }
        while (reader.NextResult());
        return values;
    }

    private static List<string> ReadFixturePath(string path, SharedStringStorageMode mode)
    {
        using var input = File.OpenRead(path);
        using var reader = ExcelReaderFactory.CreateReader(input, new ExcelReaderConfiguration { SharedStringStorageMode = mode });
        var values = new List<string>();
        while (reader.Read())
        {
            for (int c = 0; c < reader.FieldCount; c++)
                values.Add(reader.GetString(c));
        }

        return values;
    }

    private static List<string> CreateSharedStringTable() =>
        (List<string>)Activator.CreateInstance(typeof(ExcelReaderFactory).Assembly.GetType("ExcelDataReader.Core.SharedStringTable")!)!;

    private static object CreateBinarySharedStringsReader(Stream stream)
    {
        Type type = typeof(ExcelReaderFactory).Assembly.GetType("ExcelDataReader.Core.OpenXmlFormat.BinaryFormat.BiffSharedStringsReader")!;
        return type.GetConstructors()[0].Invoke([stream]);
    }

    private static void LoadBinarySharedStrings(object reader, object sink)
    {
        try
        {
            reader.GetType().GetMethod("Load")!.Invoke(reader, [sink]);
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    private static object? ReadSharedStringsRecord(object reader)
    {
        try
        {
            return reader.GetType().GetMethod("Read")!.Invoke(reader, null);
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    private static byte[] CodeUnits(string value)
    {
        byte[] bytes = new byte[value.Length * 2];
        for (int i = 0; i < value.Length; i++)
        {
            bytes[i * 2] = (byte)value[i];
            bytes[i * 2 + 1] = (byte)(value[i] >> 8);
        }

        return bytes;
    }

    private static byte[] XlsbStringItem(byte[] text, int count, byte flags = 0, int trailing = 0)
    {
        var output = new MemoryStream();
        output.WriteByte(0x13);
        for (int length = 1 + 4 + text.Length + trailing; ; length >>= 7)
        {
            output.WriteByte((byte)((length & 0x7f) | (length > 0x7f ? 0x80 : 0)));
            if (length <= 0x7f)
                break;
        }

        output.WriteByte(flags);
        output.Write(BitConverter.GetBytes(count), 0, 4);
        output.Write(text, 0, text.Length);
        output.Write(Enumerable.Repeat((byte)0xee, trailing).ToArray(), 0, trailing);
        return output.ToArray();
    }

    private sealed class Store : IDisposable
    {
        private static readonly Type Type = typeof(ExcelReaderFactory).Assembly.GetType("ExcelDataReader.Core.AutoSharedStringStore")!;
        private readonly object _store;

        public Store(SharedStringStorageMode mode, Func<string, string, Stream>? factory = null)
        {
            _store = Type.GetConstructors()[0].Invoke([new ExcelReaderConfiguration
            {
                SharedStringStorageMode = mode,
                SharedStringSpillThreshold = Budget,
                SharedStringTemporaryDirectory = TestContext.CurrentContext.WorkDirectory,
            }, factory == null ? null : new Func<string, string, Stream>((directory, kind) =>
            {
                Stream stream = factory(directory, kind);
                if (stream is FaultStream fault)
                    FactoryAction?.Invoke(fault);
                return stream;
            })]);
        }

        public Action<FaultStream>? FactoryAction { get; set; }

        public object Instance => _store;

        public int Count => (int)Property("Count");

        public bool Spilled => (bool)Property("HasSpilled");

        public long ResidentBytes => (long)Property("ResidentBytes");

        public long DiskBytes => (long)Property("DiskBytes");

        public void Add(string value) => Invoke("Add", value);

        public void Seal() => Invoke("Seal");

        public string Get(int index) => (string)Invoke("GetString", index)!;

        public void Dispose() => ((IDisposable)_store).Dispose();

        public T Field<T>(string name) => (T)Type.GetField(name, BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(_store)!;

        public object? Invoke(string method, params object[] args)
        {
            try
            {
                return Type.GetMethod(method, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(_store, args);
            }
            catch (TargetInvocationException exception) when (exception.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
                throw;
            }
        }

        private object Property(string name) => Type.GetProperty(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(_store)!;
    }

    private sealed class MappedReader : IDisposable
    {
        private static readonly Type Type = typeof(ExcelReaderFactory).Assembly.GetType("ExcelDataReader.Core.MappedSharedStringReader")!;
        private readonly object _reader;

        public MappedReader(FileStream file, long length, int viewSize)
        {
            try
            {
                _reader = Type.GetConstructors()[0].Invoke([file, length, viewSize]);
            }
            catch (TargetInvocationException exception) when (exception.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
                throw;
            }
        }

        public bool Read(long offset, byte[] buffer, int count, bool adjacent)
        {
            try
            {
                return (bool)Type.GetMethod("TryReadExactly")!.Invoke(_reader, [offset, buffer, count, adjacent])!;
            }
            catch (TargetInvocationException exception) when (exception.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
                throw;
            }
        }

        public void Dispose() => ((IDisposable)_reader).Dispose();
    }

    private sealed class FaultStream : MemoryStream
    {
        public bool FailWrite { get; set; }

        public bool FailFlush { get; set; }

        public bool FailRead { get; set; }

        public bool ShortRead { get; set; }

        public bool WasDisposed { get; private set; }

        public bool FailSeek { get; set; }

        public bool FailDispose { get; set; }

        public override long Position
        {
            get => base.Position;
            set
            {
                if (FailSeek)
                    throw new IOException("injected seek");
                base.Position = value;
            }
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            if (FailWrite)
                throw new IOException("injected write");
            base.Write(buffer, offset, count);
        }

        public override void Flush()
        {
            if (FailFlush)
                throw new IOException("injected flush");
            base.Flush();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (FailRead)
                throw new IOException("injected read");
            return base.Read(buffer, offset, ShortRead ? Math.Min(count, 3) : count);
        }

        protected override void Dispose(bool disposing)
        {
            WasDisposed = true;
            base.Dispose(disposing);
            if (FailDispose)
                throw new IOException("injected dispose");
        }
    }
}
