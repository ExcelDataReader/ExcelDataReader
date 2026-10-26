using System.Globalization;
using System.Security.Cryptography;
using ExcelDataReader.Core.OfficeCrypto;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents a BIFF stream. Records borrow a reusable frame and are valid until
/// the next read.
/// Reads must stay sequential so the RC4/XOR decryptor remains aligned.
/// </summary>
internal sealed class XlsBiffStream : IDisposable
{
    // Read-ahead buffer: a single BaseStream.Read() fills 1 024 bytes that then serve
    // ~70 sequential small records (avg 14 bytes each) without further CompoundStream calls.
    // Any Seek() sets _readAheadStart == _readAheadEnd to invalidate it.
    // Uses only byte[], Buffer.BlockCopy and Stream.Read — available on all target frameworks.
    private const int ReadAheadSize = 1024;

    private readonly byte[] _headerBuffer = new byte[4];

    private readonly byte[] _readAheadBuffer = new byte[ReadAheadSize];
    private BiffFrameBuffer _frameBuffer = new(64);
    private int _readAheadStart;
    private int _readAheadEnd;

    // Reused decryption buffers: one 1024-byte encryption block in/out, and the current block key. Allocated on first use.
    private byte[]? _decryptOutputBuffer;
    private byte[]? _blockKeyBuffer;

    public XlsBiffStream(Stream baseStream, int offset = 0)
    {
        BaseStream = baseStream;
        CipherBlock = -1;
        Position = offset;
    }

    /// <summary>
    /// Gets the size of BIFF stream in bytes.
    /// </summary>
    public int Size => (int)BaseStream.Length;

    /// <summary>
    /// Gets or sets the current position in BIFF stream.
    /// </summary>
    public int Position
    {
        get => (int)BaseStream.Position - (_readAheadEnd - _readAheadStart);
        set => Seek(value, SeekOrigin.Begin);
    }

    public Stream BaseStream { get; }

    public byte[]? SecretKey { get; private set; }

    public EncryptionInfo? Encryption { get; private set; }

    public SymmetricAlgorithm? Cipher { get; private set; }

    /// <summary>
    /// Gets or sets the ICryptoTransform instance used to decrypt the current block.
    /// </summary>
    public ICryptoTransform? CipherTransform { get; set; }

    /// <summary>
    /// Gets or sets the current block number being decrypted with CipherTransform.
    /// </summary>
    public int CipherBlock { get; set; }

    /// <summary>
    /// Sets stream pointer to the specified offset.
    /// </summary>
    /// <param name="offset">Offset value.</param>
    /// <param name="origin">Offset origin.</param>
    public void Seek(int offset, SeekOrigin origin)
    {
        // Discard buffered bytes: they belong to the old position.
        _readAheadStart = 0;
        _readAheadEnd = 0;
        BaseStream.Seek(offset, origin);

        if (Position < 0)
            throw new ArgumentOutOfRangeException(string.Format(CultureInfo.InvariantCulture, "{0} On offset={1}", Errors.ErrorBiffIlegalBefore, offset));
        if (Position > Size)
            throw new ArgumentOutOfRangeException(string.Format(CultureInfo.InvariantCulture, "{0} On offset={1}", Errors.ErrorBiffIlegalAfter, offset));

        if (SecretKey != null)
        {
            CreateBlockDecryptor(offset / 1024);
            AlignBlockDecryptor(offset % 1024);
        }
    }

    // All consumers must finish reading a frame before requesting the next one.
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    public bool ReadFrame(out XlsBiffFrame result)
    {
        result = default;
        int streamSize = Size;
        if ((uint)Position + 4 >= streamSize)
            return false;

        ReadFrameHeader(out var recordOffset, out var id, out var recordSize);
        int size = 4 + recordSize;
        byte[] buffer = _frameBuffer.GetBuffer(size);
        ReadFramePayload(recordOffset, id, recordSize, buffer);
        if (Position <= Size)
        {
            result = new XlsBiffFrame(buffer);
            return true;
        }

        return false;
    }

    public void Dispose()
    {
        _frameBuffer.Clear();
        CipherTransform?.Dispose();
        Cipher?.Dispose();
    }

    internal void InitializeEncryption(byte[]? secretKey, EncryptionInfo? encryption)
    {
        SecretKey = secretKey;
        Encryption = encryption;
        Cipher = encryption?.CreateCipher();
    }

    /// <summary>
    /// Decrypts <paramref name="count"/> bytes starting at <paramref name="inputOffset"/>. The BIFF RC4 and XOR
    /// transforms are stream ciphers that process any count in TransformBlock, so call them directly instead of
    /// allocating a CryptoStream per chunk.
    /// </summary>
    private static void Transform(ICryptoTransform transform, byte[] input, int inputOffset, int count, byte[] output)
    {
        if (transform is RC4Managed.RC4Transform or XorManaged.XorTransform)
            transform.TransformBlock(input, inputOffset, count, output, 0);
        else
            CryptoHelpers.DecryptBytes(transform, input, inputOffset, count, output);
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private void ReadFrameHeader(out int offset, out BIFFRECORDTYPE id, out ushort size)
    {
        offset = Position;
        ReadFromReadAhead(_headerBuffer, 0, 4);
        id = (BIFFRECORDTYPE)BitConverter.ToUInt16(_headerBuffer, 0);
        size = BitConverter.ToUInt16(_headerBuffer, 2);
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private void ReadFramePayload(int offset, BIFFRECORDTYPE id, ushort size, byte[] bytes)
    {
        Array.Copy(_headerBuffer, bytes, 4);
        ReadFromReadAhead(bytes, 4, size);
        if (SecretKey != null)
            DecryptRecord(offset, id, bytes, 4 + size);
    }

    // Serves 'count' bytes into dest[destOffset..], draining the read-ahead buffer first.
    // Refill read-ahead in one call when possible, accumulating short reads as needed.
    // For large bodies that exceed the remaining buffer the tail is read directly from
    // BaseStream (avoids per-1024-chunk overhead for SST / other big records).
    private void ReadFromReadAhead(byte[] dest, int destOffset, int count)
    {
        // Fast path: serve from the in-memory buffer when possible.
        int available = _readAheadEnd - _readAheadStart;
        if (available > 0)
        {
            int fromBuffer = Math.Min(available, count);
            Buffer.BlockCopy(_readAheadBuffer, _readAheadStart, dest, destOffset, fromBuffer);
            _readAheadStart += fromBuffer;

            if (fromBuffer == count)
                return;

            destOffset += fromBuffer;
            count -= fromBuffer;
        }

        // Buffer was exhausted before satisfying the request.
        // For small leftovers refill the buffer and serve from it;
        // for large reads go directly to BaseStream to avoid per-chunk overhead.
        if (count <= ReadAheadSize)
        {
            _readAheadStart = 0;
            _readAheadEnd = BaseStream.Read(_readAheadBuffer, 0, ReadAheadSize);
            while (_readAheadEnd < count)
            {
                int read = BaseStream.Read(_readAheadBuffer, _readAheadEnd, ReadAheadSize - _readAheadEnd);
                if (read == 0)
                    throw new EndOfStreamException();
                _readAheadEnd += read;
            }

            Buffer.BlockCopy(_readAheadBuffer, 0, dest, destOffset, count);
            _readAheadStart = count;
        }
        else
        {
            // Large record body: read directly into the caller's buffer.
            // The read-ahead buffer remains empty (_readAheadStart == _readAheadEnd == 0)
            // and will be refilled on the next call.
            _readAheadStart = 0;
            _readAheadEnd = 0;
            if (BaseStream.ReadAtLeast(dest, destOffset, count, count, throwOnEndOfStream: false) != count)
                throw new EndOfStreamException();
        }
    }

    /// <summary>
    /// Create or re-key the ICryptoTransform instance used to decrypt a 1024-byte block.
    /// </summary>
    private void CreateBlockDecryptor(int blockNumber)
    {
        if (Encryption == null || SecretKey == null || Cipher == null)
            throw new InvalidOperationException("Encryption is not initialized.");

        // This runs for every 1024-byte block and on every seek, so re-key the existing
        // RC4/XOR transform in place instead of allocating a new key and transform.
        switch (CipherTransform)
        {
            case RC4Managed.RC4Transform rc4:
                var blockKey = _blockKeyBuffer ??= new byte[CryptoHelpers.MaxHashSize];
                var keyLength = Encryption.GenerateBlockKey(blockNumber, SecretKey, blockKey);
                rc4.Reset(blockKey, keyLength);
                break;
            case XorManaged.XorTransform xor:
                // The XOR key doesn't depend on the block; the array index is set per record.
                xor.XorArrayIndex = 0;
                break;
            default:
                CipherTransform?.Dispose();
                CipherTransform = Cipher.CreateDecryptor(Encryption.GenerateBlockKey(blockNumber, SecretKey), null);
                break;
        }

        CipherBlock = blockNumber;
    }

    /// <summary>
    /// Decrypt some dummy bytes to align the decryptor with the position in the current 1024-byte block.
    /// </summary>
    private void AlignBlockDecryptor(int blockOffset)
    {
        var cipherTransform = CipherTransform ?? throw new InvalidOperationException("Decryptor is not initialized.");

        var bytes = _decryptOutputBuffer ??= new byte[1024];
        Transform(cipherTransform, bytes, 0, blockOffset, bytes);
    }

    private void DecryptRecord(int startPosition, BIFFRECORDTYPE id, byte[] bytes, int recordSize)
    {
        var encryption = Encryption ?? throw new InvalidOperationException("Encryption is not initialized.");

        // Decrypt the last read record, find it's start offset relative to the current stream position
        int startDecrypt = 4;
        switch (id)
        {
            case BIFFRECORDTYPE.BOF:
            case BIFFRECORDTYPE.FILEPASS:
            case BIFFRECORDTYPE.INTERFACEHDR:
                startDecrypt = recordSize;
                break;
            case BIFFRECORDTYPE.BOUNDSHEET:
                startDecrypt += 4; // For some reason the sheet offset is not encrypted
                break;
        }

        // Max chunk size per iteration is 1024 (one encryption block boundary).
        var outputBlock = _decryptOutputBuffer ??= new byte[1024];
        var position = 0;
        while (position < recordSize)
        {
            var offset = startPosition + position;
            int blockNumber = offset / 1024;
            var blockOffset = offset % 1024;

            if (blockNumber != CipherBlock)
            {
                CreateBlockDecryptor(blockNumber);
            }

            if (encryption.IsXor)
            {
                // Bypass everything and hook into the XorTransform instance to set the XorArrayIndex pr record.
                // This is a hack to use the XorTransform otherwise transparently to the other encryption methods.
                var xorTransform = CipherTransform as XorManaged.XorTransform;
                if (xorTransform == null)
                    throw new InvalidOperationException("XOR decryptor is not initialized.");
                xorTransform.XorArrayIndex = offset + recordSize - 4;
            }

            // Decrypt at most up to the next 1024 byte boundary
            var chunkSize = Math.Min(recordSize - position, 1024 - blockOffset);

            var cipherTransform = CipherTransform ?? throw new InvalidOperationException("Decryptor is not initialized.");
            Transform(cipherTransform, bytes, position, chunkSize, outputBlock);

            // Bytes before startDecrypt must still pass through the cipher to advance its
            // keystream, but keep their original value.
            var copyFrom = Math.Max(position, startDecrypt);
            if (copyFrom < position + chunkSize)
                Buffer.BlockCopy(outputBlock, copyFrom - position, bytes, copyFrom, position + chunkSize - copyFrom);

            position += chunkSize;
        }
    }
}