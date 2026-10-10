using System.Text;

namespace ExcelDataReader.Core;

internal sealed class AutoSharedStringStore : ISharedStringStore
{
    private const int ChunkSize = 32768;
    private const int IndexBufferSize = 4096;
    private const int DiskEntrySize = 16;
    private const long FixedBytes = 256;
    private const long DiskBufferBytes = ChunkSize + IndexBufferSize + 64L + 2 * (4096L + 256);
    private readonly long _limit;
    private readonly long _cacheLimit;
    private readonly string? _directory;
    private readonly Func<string, string, Stream> _createFile;
    private List<string>? _strings = [];
    private byte[]? _buffer;
    private byte[]? _entryBuffer;
    private Stream? _payload;
    private Stream? _index;
    private MappedSharedStringReader? _payloadMapping;
    private MappedSharedStringReader? _indexMapping;
    private long _length;
    private long _payloadWindowOffset;
    private int _payloadWindowLength;
    private long _indexWindowOffset;
    private int _indexWindowLength;
    private long _previousPayloadOffset = -1;
    private long _previousPayloadEnd = -1;
    private int _previousIndex = -2;
    private long _allocatedBytes = FixedBytes;
    private long _cacheBytes;
    private long _lastValueBytes;
    private int[]? _cacheIndices;
    private int[]? _cacheSeen;
    private string?[]? _cacheValues;
    private int _evictionCursor;
    private int _cachedIndex = -1;
    private string? _cachedValue;
    private bool _sealed;
    private bool _disposed;

    public AutoSharedStringStore(ExcelReaderConfiguration configuration, Func<string, string, Stream>? createFile = null)
    {
        if (configuration.SharedStringStorageMode != SharedStringStorageMode.SpillToDisk)
            throw new ArgumentOutOfRangeException(nameof(configuration), "Unknown shared string storage mode.");
        _limit = configuration.SharedStringSpillThreshold;
        if (_limit < 1024 * 1024)
            throw new ArgumentOutOfRangeException(nameof(configuration), "SharedStringSpillThreshold must be at least 1 MiB.");
        _cacheLimit = Math.Min(_limit / 8, 8L * 1024 * 1024);
        _directory = configuration.SharedStringTemporaryDirectory;
        _createFile = createFile ?? CreateFile;
    }

    public int Count { get; private set; }

    internal long ResidentBytes => _allocatedBytes + _cacheBytes;

    internal bool HasSpilled => _payload != null;

    internal long DiskBytes => _payload == null ? 0 : checked(_length + (long)Count * DiskEntrySize);

    private long CacheBudget => Math.Min(_cacheLimit, _limit - _allocatedBytes);

    public void Add(string value)
    {
        CheckCanAdd();
        if (TryAddNormal(value, ValueBytes(value)))
            return;
        WriteString(value);
        Count++;
    }

    public void AddUtf16(byte[] buffer, int offset, int characterCount)
    {
        CheckCanAdd();
        if (offset < 0 || characterCount < 0 || characterCount > (buffer.Length - offset) / 2)
            throw new ArgumentOutOfRangeException(nameof(characterCount));
        if (TryAddUtf16(buffer, offset, characterCount))
            return;
        if (!TryWriteUtf16(buffer, offset, characterCount))
            WriteString(Encoding.Unicode.GetString(buffer, offset, characterCount * 2));
        Count++;
    }

    public void Seal()
    {
        ThrowIfDisposed();
        if (_sealed)
            return;
        try
        {
            CompleteWrites();
        }
        catch (Exception exception)
        {
            ResourceCleanup.DisposeAll(exception, this);
            throw;
        }

        _sealed = true;
    }

    public string GetString(int index)
    {
        ThrowIfDisposed();
        if (!_sealed)
            throw new InvalidOperationException("Shared string storage is not sealed.");
        if ((uint)index >= (uint)Count)
            throw new ArgumentOutOfRangeException(nameof(index));
        if (_strings != null)
            return _strings[index];

        if (_cachedIndex == index)
            return _cachedValue!;
        if (_cacheIndices != null)
        {
            int slot = index & (_cacheIndices.Length - 1);
            if (_cacheIndices[slot] == index)
            {
                string cached = _cacheValues![slot]!;
                CacheLast(index, cached);
                return cached;
            }
        }

        Entry entry;
        {
            long offset = checked((long)index * DiskEntrySize);
            bool adjacent = Math.Abs((long)index - _previousIndex) == 1;
            int start = ReadWindow(_index!, _indexMapping, _entryBuffer!, offset, DiskEntrySize, (long)Count * DiskEntrySize, adjacent, ref _indexWindowOffset, ref _indexWindowLength);
            _previousIndex = index;
            byte[] indexBuffer = _entryBuffer!;
            entry = new Entry(ReadInt64(indexBuffer, start), ReadInt32(indexBuffer, start + 8));
            long end = checked(entry.Offset + (long)entry.Length * 2);
            if (entry.Offset < 0 || (entry.Offset & 1) != 0 || entry.Length < 0 || end > _length)
                throw new IOException("Invalid shared string index data.");
        }

        string value;
        if (entry.Length == 0)
        {
            value = string.Empty;
        }
        else
        {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
            value = string.Create(entry.Length, (Store: this, Entry: entry), static (chars, state) =>
            {
                state.Store.CopyCharacters(state.Entry, chars);
            });
#else
            var chars = new char[entry.Length];
            CopyCharacters(entry, chars);
            value = new string(chars);
#endif
        }

        CacheDecoded(index, value);

        return value;
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        try
        {
            ResourceCleanup.DisposeAll(null, _payloadMapping, _indexMapping, _payload, _index);
        }
        finally
        {
            _payload = null;
            _index = null;
            _payloadMapping = null;
            _indexMapping = null;
            _strings = null;
            _buffer = null;
            _entryBuffer = null;
            _cachedValue = null;
            _cacheIndices = null;
            _cacheSeen = null;
            _cacheValues = null;
            _lastValueBytes = 0;
            _cacheBytes = 0;
            _allocatedBytes = FixedBytes;
        }
    }

    public void Reserve(int capacity)
    {
        CheckCanAdd();
        if (_strings == null || capacity <= _strings.Capacity)
            return;
        long growth = (long)(capacity - _strings.Capacity) * IntPtr.Size;
        if (checked(_allocatedBytes + growth) > _limit)
            return;
        _strings.Capacity = capacity;
        _allocatedBytes += growth;
    }

    private static long ValueBytes(string value) => StringBytes(value.Length);

    private static long StringBytes(int length) => length == 0 ? 0 : 32L + (long)length * 2;

    private static int Capacity(int current, int required)
    {
        if (current >= required)
            return current;
        return (int)Math.Min(int.MaxValue, Math.Max((long)required, current == 0 ? 4 : (long)current * 2));
    }

    private static FileStream CreateFile(string directory, string kind)
    {
        string path = Path.Combine(directory, "ExcelDataReader-sst-" + Guid.NewGuid().ToString("N") + "-" + kind + ".tmp");
        return new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose);
    }

    private static void ReadExactly(Stream stream, byte[] buffer, int count)
    {
        int offset = 0;
        while (offset < count)
        {
            int read = stream.Read(buffer, offset, count - offset);
            if (read == 0)
                throw new EndOfStreamException("Unexpected end of shared string storage.");
            offset += read;
        }
    }

    private static int ReadInt32(byte[] bytes, int offset) =>
        bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24);

    private static long ReadInt64(byte[] bytes, int offset) =>
        (uint)ReadInt32(bytes, offset) | ((long)ReadInt32(bytes, offset + 4) << 32);

    private static int ReadWindow(Stream stream, MappedSharedStringReader? mapping, byte[] buffer, long offset, int count, long length, bool adjacent, ref long windowOffset, ref int windowLength)
    {
        if (offset < windowOffset || offset + count > windowOffset + windowLength)
        {
            // Read ahead only for adjacent entries; random misses keep their exact read size.
            windowLength = 0;
            windowOffset = adjacent ? offset / buffer.Length * buffer.Length : offset;
            int read = adjacent ? (int)Math.Min(buffer.Length, length - windowOffset) : count;
            if (adjacent || mapping == null || !mapping.TryReadExactly(windowOffset, buffer, read, adjacent))
            {
                stream.Position = windowOffset;
                ReadExactly(stream, buffer, read);
            }

            windowLength = read;
        }

        return (int)(offset - windowOffset);
    }

    private bool TryAddNormal(string value, long bytes)
    {
        if (!FitsNormal(bytes, out int capacity, out long growth))
            return false;
        AddNormal(value, capacity, growth);
        return true;
    }

    private bool TryAddUtf16(byte[] buffer, int offset, int characterCount)
    {
        // The budget is checked before the string is allocated.
        if (!FitsNormal(StringBytes(characterCount), out int capacity, out long growth))
            return false;
        string value = Encoding.Unicode.GetString(buffer, offset, characterCount * 2);
        if (value.Length != characterCount)
            return TryAddNormal(value, ValueBytes(value));
        AddNormal(value, capacity, growth);
        return true;
    }

    private bool FitsNormal(long bytes, out int capacity, out long growth)
    {
        capacity = 0;
        growth = 0;
        if (_strings == null)
            return false;
        capacity = Capacity(_strings.Capacity, Count + 1);
        growth = checked((long)(capacity - _strings.Capacity) * IntPtr.Size + bytes);
        if (checked(_allocatedBytes + growth) > _limit)
        {
            Spill();
            return false;
        }

        return true;
    }

    private void AddNormal(string value, int capacity, long growth)
    {
        _strings!.Capacity = capacity;
        _strings.Add(value);
        _allocatedBytes += growth;
        Count++;
    }

    private void WriteString(string value)
    {
        Entry entry = new(_length, value.Length);
        byte[] buffer = _buffer!;
        int used = 0;
        foreach (char c in value)
        {
            buffer[used++] = (byte)c;
            buffer[used++] = (byte)(c >> 8);
            if (used == buffer.Length)
            {
                _payload!.Write(buffer, 0, used);
                used = 0;
            }
        }

        if (used > 0)
            _payload!.Write(buffer, 0, used);
        WriteEntry(entry);
        _length = checked(entry.Offset + (long)entry.Length * 2);
    }

    private bool TryWriteUtf16(byte[] bytes, int offset, int characterCount)
    {
        int end = offset + characterCount * 2;
        for (int i = offset; i < end; i += 2)
        {
            char c = (char)(bytes[i] | (bytes[i + 1] << 8));
            if (!char.IsSurrogate(c))
                continue;
            if (char.IsHighSurrogate(c) && i + 3 < end &&
                char.IsLowSurrogate((char)(bytes[i + 2] | (bytes[i + 3] << 8))))
            {
                i += 2;
                continue;
            }

            // Malformed UTF-16 is decoded by the caller so replacement characters match normal decoding.
            return false;
        }

        if (characterCount != 0)
            _payload!.Write(bytes, offset, characterCount * 2);
        Entry entry = new(_length, characterCount);
        WriteEntry(entry);
        _length = checked(entry.Offset + (long)entry.Length * 2);
        return true;
    }

    private void CacheLast(int index, string value)
    {
        _cacheBytes -= _lastValueBytes;
        _lastValueBytes = 0;
        _cachedIndex = -1;
        _cachedValue = null;
        long bytes = ValueBytes(value);
        if (bytes > CacheBudget)
            return;
        if (_cacheIndices != null && (long)_cacheIndices.Length * 16 + 96 + bytes > CacheBudget)
        {
            _cacheIndices = null;
            _cacheSeen = null;
            _cacheValues = null;
            _cacheBytes = 0;
            _evictionCursor = 0;
        }

        EvictUntil(bytes);
        _cachedIndex = index;
        _cachedValue = value;
        _lastValueBytes = bytes;
        _cacheBytes += bytes;
    }

    private void CacheDecoded(int index, string value)
    {
        CacheLast(index, value);
        if (_cacheIndices == null)
        {
            long available = CacheBudget - _cacheBytes;
            int slots = 16;
            while (slots < 4096 && (long)slots * 2 * 16 + 96 <= available / 2)
                slots *= 2;
            long metadata = (long)slots * 16 + 96;
            if (metadata > available / 2)
                return;
            _cacheIndices = new int[slots];
            _cacheSeen = new int[slots];
            _cacheValues = new string?[slots];
            for (int i = 0; i < slots; i++)
            {
                _cacheIndices[i] = -1;
                _cacheSeen[i] = -1;
            }

            _cacheBytes += metadata;
        }

        int slot = index & (_cacheIndices.Length - 1);

        // A one-off scan may update admission history, but cannot displace cached values.
        bool repeated = _cacheSeen![slot] == index;
        _cacheSeen[slot] = index;
        long bytes = ValueBytes(value);
        if (!repeated || bytes > CacheBudget / 4)
            return;
        EvictSlot(slot);
        long metadataBytes = (long)_cacheIndices.Length * 16 + 96;
        if (metadataBytes + _lastValueBytes + bytes > CacheBudget)
            return;
        EvictUntil(bytes);
        _cacheIndices[slot] = index;
        _cacheValues![slot] = value;
        _cacheBytes += bytes;
    }

    private void EvictUntil(long bytes)
    {
        while (_cacheBytes + bytes > CacheBudget && _cacheIndices != null)
        {
            EvictSlot(_evictionCursor);
            _evictionCursor = (_evictionCursor + 1) & (_cacheIndices.Length - 1);
        }
    }

    private void EvictSlot(int slot)
    {
        if (_cacheIndices![slot] < 0)
            return;
        _cacheBytes -= ValueBytes(_cacheValues![slot]!);
        _cacheIndices[slot] = -1;
        _cacheValues[slot] = null;
    }

    private void Spill()
    {
        string directory = _directory ?? Path.GetTempPath();
        try
        {
            _payload = _createFile(directory, "payload");
            _index = _createFile(directory, "index");
            _buffer = new byte[ChunkSize];
            _entryBuffer = new byte[IndexBufferSize];
            for (int i = 0; i < Count; i++)
                WriteString(_strings![i]);

            _strings = null;
            _allocatedBytes = FixedBytes + DiskBufferBytes;
        }
        catch (Exception exception)
        {
            ResourceCleanup.DisposeAll(exception, this);
            throw;
        }
    }

    private void CompleteWrites()
    {
        _payload?.Flush();
        _index?.Flush();
        if (_length != 0 && _payload is FileStream payload)
        {
            _payloadMapping = new MappedSharedStringReader(payload, _length);
            _allocatedBytes += MappedSharedStringReader.ResidentBytes;
        }

        if (Count != 0 && _index is FileStream index)
        {
            _indexMapping = new MappedSharedStringReader(index, (long)Count * DiskEntrySize);
            _allocatedBytes += MappedSharedStringReader.ResidentBytes;
        }
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private void CheckCanAdd()
    {
        ThrowIfDisposed();
        if (_sealed)
            throw new InvalidOperationException("Shared string storage is sealed.");
        if (Count == int.MaxValue)
            throw new IOException("Too many shared strings.");
    }

    private void WriteEntry(Entry entry)
    {
        byte[] indexBuffer = _entryBuffer!;
        long offset = entry.Offset;
        for (int i = 0; i < 8; i++)
        {
            indexBuffer[i] = (byte)offset;
            offset >>= 8;
        }

        int length = entry.Length;
        for (int i = 8; i < 12; i++)
        {
            indexBuffer[i] = (byte)length;
            length >>= 8;
        }

        _index!.Write(indexBuffer, 0, DiskEntrySize);
    }

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
    private void CopyCharacters(Entry entry, Span<char> chars)
#else
    private void CopyCharacters(Entry entry, char[] chars)
#endif
    {
        long end = entry.Offset + (long)entry.Length * 2;
        bool adjacent = entry.Offset == _previousPayloadEnd || end == _previousPayloadOffset;
        for (int start = 0; start < chars.Length;)
        {
            long offset = entry.Offset + (long)start * 2;
            int bytes = (int)Math.Min((long)(chars.Length - start) * 2, _buffer!.Length);
            if (adjacent)
                bytes = Math.Min(bytes, _buffer.Length - (int)(offset % _buffer.Length));
            int bufferStart = ReadWindow(_payload!, _payloadMapping, _buffer, offset, bytes, _length, adjacent, ref _payloadWindowOffset, ref _payloadWindowLength);
            int count = bytes / 2;
            for (int i = 0; i < count; i++)
                chars[start + i] = (char)(_buffer[bufferStart + i * 2] | (_buffer[bufferStart + i * 2 + 1] << 8));
            start += count;
        }

        _previousPayloadOffset = entry.Offset;
        _previousPayloadEnd = end;
    }

    private void ThrowIfDisposed()
    {
#if NET8_0_OR_GREATER
        ObjectDisposedException.ThrowIf(_disposed, this);
#else
        if (_disposed)
            throw new ObjectDisposedException(nameof(AutoSharedStringStore));
#endif
    }

    private readonly struct Entry(long offset, int length)
    {
        public long Offset { get; } = offset;

        public int Length { get; } = length;
    }
}
