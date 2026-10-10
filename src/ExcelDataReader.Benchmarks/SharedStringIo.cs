using System.Diagnostics;
using System.Reflection;
using System.Text;
#if NET8_0_OR_GREATER
using System.Text.Json;
#endif
using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Benchmarks;

internal static class SharedStringIo
{
    public static void Run()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        string library = Environment.GetEnvironmentVariable("EDR_SST_LIBRARY");
        Assembly assembly = string.IsNullOrEmpty(library) ? typeof(ExcelReaderFactory).Assembly : Assembly.LoadFile(Path.GetFullPath(library));
        foreach (string backend in new[] { "Memory", "File" })
        {
            foreach (string pattern in new[] { "Sequential", "Permuted", "Random", "Hot" })
            {
                for (int iteration = -3; iteration < 5; iteration++)
                    Measure(assembly, backend, pattern, iteration);
            }
        }
    }

    private static void Measure(Assembly assembly, string backend, string pattern, int iteration)
    {
        const int count = 100000;
        const int length = 64;
        string[] values = Enumerable.Range(0, count).Select(i => SharedStringWorkbook.Value(i, length, pattern == "Hot")).ToArray();
        Type configurationType = assembly.GetType("ExcelDataReader.ExcelReaderConfiguration")!;
        object configuration = Activator.CreateInstance(configurationType)!;
        configurationType.GetProperty("SharedStringStorageMode")!.SetValue(configuration, Enum.Parse(assembly.GetType("ExcelDataReader.SharedStringStorageMode")!, "SpillToDisk"));
        configurationType.GetProperty("SharedStringSpillThreshold")!.SetValue(configuration, 1024L * 1024);
        configurationType.GetProperty("SharedStringTemporaryDirectory")!.SetValue(configuration, Environment.CurrentDirectory);
        var streams = new Dictionary<string, CountingStream>();
        Func<string, string, Stream> factory = (directory, kind) =>
        {
            Stream stream = backend == "Memory" ? new MemoryStream() :
                new FileStream(Path.Combine(directory, "ExcelDataReader-io-" + Guid.NewGuid().ToString("N") + "-" + kind), FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose);
            var counted = new CountingStream(stream);
            streams.Add(kind, counted);
            return counted;
        };
        Type storeType = assembly.GetType("ExcelDataReader.Core.AutoSharedStringStore")!;
        object store = storeType.GetConstructors()[0].Invoke([configuration, factory]);
        using var disposable = (IDisposable)store;
        var add = (Action<string>)storeType.GetMethod("Add")!.CreateDelegate(typeof(Action<string>), store);
        var seal = (Action)storeType.GetMethod("Seal")!.CreateDelegate(typeof(Action), store);
        var get = (Func<int, string>)storeType.GetMethod("GetString")!.CreateDelegate(typeof(Func<int, string>), store);
        Func<int, string> decoded = index => values[index];
        var timer = Stopwatch.StartNew();
        ulong expected = ReadValues(decoded, values, pattern);
        double decodedLookupMs = timer.Elapsed.TotalMilliseconds;
        timer.Restart();
        foreach (string value in values)
            add(value);
        seal();
        double constructionMs = timer.Elapsed.TotalMilliseconds;
        var constructionIo = streams.ToDictionary(pair => pair.Key, pair => pair.Value.Snapshot());
        foreach (CountingStream stream in streams.Values)
            stream.Reset();
        timer.Restart();
        int references = SharedStringWorkbook.References(count, pattern);
        ulong hash = ReadValues(get, values, pattern);
        double lookupMs = timer.Elapsed.TotalMilliseconds;
        if (hash != expected)
            throw new InvalidOperationException("Direct store checksum mismatch.");
        var lookupIo = streams.ToDictionary(pair => pair.Key, pair => pair.Value.Snapshot());
        long residentBytes = (long)storeType.GetProperty("ResidentBytes", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(store)!;
        if (iteration < 0)
            return;
        var result = new { backend, pattern, iteration, count, length, references, constructionMs, decodedLookupMs, lookupMs, hash, residentBytes, constructionIo, lookupIo };
#if NET8_0_OR_GREATER
        Console.WriteLine(JsonSerializer.Serialize(result));
#else
        Console.WriteLine(string.Join("; ", result.GetType().GetProperties().Select(p => p.Name + "=" + p.GetValue(result))));
        foreach (string kind in streams.Keys)
            Console.WriteLine(kind + ": construction " + constructionIo[kind] + "; lookup " + lookupIo[kind]);
#endif
    }

    private static ulong ReadValues(Func<int, string> get, string[] values, string pattern)
    {
        ulong hash = 14695981039346656037;
        int references = SharedStringWorkbook.References(values.Length, pattern);
        for (int position = 0; position < references; position++)
        {
            int index = SharedStringWorkbook.Reference(position, values.Length, pattern);
            string actual = get(index);
            if (actual != values[index])
                throw new InvalidOperationException("Direct store changed string contents.");
            hash = SharedStringWorkbook.Hash(hash, actual);
        }

        return hash;
    }

    private sealed class CountingStream(Stream inner) : Stream
    {
        private long _reads;
        private long _readBytes;
        private long _writes;
        private long _writeBytes;
        private long _seeks;
        private int _largestRead;

        public override bool CanRead => inner.CanRead;

        public override bool CanSeek => inner.CanSeek;

        public override bool CanWrite => inner.CanWrite;

        public override long Length => inner.Length;

        public override long Position
        {
            get => inner.Position;
            set
            {
                _seeks++;
                inner.Position = value;
            }
        }

        public object Snapshot() => new { Reads = _reads, ReadBytes = _readBytes, Writes = _writes, WriteBytes = _writeBytes, Seeks = _seeks, LargestRead = _largestRead };

        public void Reset()
        {
            _reads = _readBytes = _writes = _writeBytes = _seeks = 0;
            _largestRead = 0;
        }

        public override void Flush() => inner.Flush();

        public override int Read(byte[] buffer, int offset, int count)
        {
            _reads++;
            _largestRead = Math.Max(_largestRead, count);
            int read = inner.Read(buffer, offset, count);
            _readBytes += read;
            return read;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            _seeks++;
            return inner.Seek(offset, origin);
        }

        public override void SetLength(long value) => inner.SetLength(value);

        public override void Write(byte[] buffer, int offset, int count)
        {
            _writes++;
            _writeBytes += count;
            inner.Write(buffer, offset, count);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
