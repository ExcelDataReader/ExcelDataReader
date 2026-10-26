namespace ExcelDataReader.Core.OfficeCrypto;

/// <summary>
/// A seekable stream for reading an EncryptedPackage blob using OpenXml Agile Encryption. 
/// </summary>
internal sealed class AgileEncryptedPackageStream : Stream
{
    private const int SegmentLength = 4096;

    private System.Security.Cryptography.SymmetricAlgorithm? _cipher;
    private byte[]? _decryptBuffer;
    private byte[]? _ivBuffer;
    private long _position;
    private int _loadedSegment = -1;

    public AgileEncryptedPackageStream(Stream stream, byte[] key, byte[] iv, EncryptionInfo encryption)
    {
        Stream = stream;
        Key = key;
        IV = iv;
        Encryption = encryption;

        if (Stream.ReadAtLeast(SegmentBytes, 0, 8, 8, throwOnEndOfStream: false) != 8)
            throw new EndOfStreamException("Incomplete encrypted package header.");
        DecryptedLength = BitConverter.ToInt64(SegmentBytes, 0);
        if (DecryptedLength < 0)
            throw new System.Security.Cryptography.CryptographicException("Invalid encrypted package length.");
        if (DecryptedLength > 0)
        {
            try
            {
                ReadSegment(0);
            }
            catch (Exception error)
            {
                ResourceCleanup.DisposeAll(error, _cipher);
                throw;
            }
        }
    }

    public override bool CanRead => true;

    public override bool CanSeek => true;

    public override bool CanWrite => false;

    public override long Length => DecryptedLength;

    public override long Position { get => _position; set => Seek(value, SeekOrigin.Begin); }

    private Stream? Stream { get; set; }

    private byte[] Key { get; }

    private byte[] IV { get; }

    private EncryptionInfo Encryption { get; }

    private byte[] SegmentBytes { get; set; } = new byte[SegmentLength];

    private long DecryptedLength { get; }

    public override void Flush()
    {
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        return Read(buffer.AsSpan(offset, count));
#else
        ArgumentNullException.ThrowIfNull(buffer);
        if (offset < 0 || count < 0 || offset > buffer.Length - count)
            throw new ArgumentOutOfRangeException(nameof(offset));
        _ = Stream ?? throw new ObjectDisposedException(nameof(AgileEncryptedPackageStream));
        count = (int)Math.Min(count, Math.Max(0, Length - Position));

        int index = 0;
        while (index < count)
        {
            int segment = checked((int)(_position / SegmentLength));
            int segmentOffset = (int)(_position % SegmentLength);
            if (_loadedSegment != segment)
                ReadSegment(segment);
            var chunkSize = Math.Min(count - index, SegmentLength - segmentOffset);
            Array.Copy(SegmentBytes, segmentOffset, buffer, offset + index, chunkSize);
            index += chunkSize;
            _position += chunkSize;
        }

        return index;
#endif
    }

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
    public override int Read(Span<byte> buffer)
    {
        _ = Stream ?? throw new ObjectDisposedException(nameof(AgileEncryptedPackageStream));
        int count = (int)Math.Min(buffer.Length, Math.Max(0, Length - Position));

        int index = 0;
        while (index < count)
        {
            int segment = checked((int)(_position / SegmentLength));
            int segmentOffset = (int)(_position % SegmentLength);
            if (_loadedSegment != segment)
                ReadSegment(segment);
            var chunkSize = Math.Min(count - index, SegmentLength - segmentOffset);
            SegmentBytes.AsSpan(segmentOffset, chunkSize).CopyTo(buffer[index..]);
            index += chunkSize;
            _position += chunkSize;
        }

        return index;
    }
#endif

    public override long Seek(long offset, SeekOrigin origin)
    {
        _ = Stream ?? throw new ObjectDisposedException(nameof(AgileEncryptedPackageStream));
        long position = origin switch
        {
            SeekOrigin.Begin => offset,
            SeekOrigin.Current => checked(Position + offset),
            SeekOrigin.End => checked(Length + offset),
            _ => throw new ArgumentOutOfRangeException(nameof(origin)),
        };
        if (position < 0)
            throw new IOException("Cannot seek before the beginning of the encrypted stream.");
        _position = position;
        return position;
    }

    public override void SetLength(long value)
    {
        throw new NotSupportedException();
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        throw new NotSupportedException();
    }

    protected override void Dispose(bool disposing)
    {
        try
        {
            if (disposing)
            {
                var stream = Stream;
                var cipher = _cipher;
                Stream = null;
                _cipher = null;
                ResourceCleanup.DisposeAll(null, stream, cipher);
            }
        }
        finally
        {
            base.Dispose(disposing);
        }
    }

    private void ReadSegment(int segmentIndex)
    {
        var stream = Stream ?? throw new ObjectDisposedException(nameof(AgileEncryptedPackageStream));

        // NOTE: +8 skips EncryptedPackage header
        long offset = (long)segmentIndex * SegmentLength;
        _loadedSegment = -1;
        stream.Seek(8 + offset, SeekOrigin.Begin);

        if (_cipher == null)
        {
            _cipher = Encryption.CreateCipher();
            _cipher.Key = Key;
            _cipher.Padding = System.Security.Cryptography.PaddingMode.None;
        }

        int plaintextLength = (int)Math.Min(SegmentLength, Length - offset);
        int blockSize = _cipher.BlockSize / 8;
        int encryptedLength = ((plaintextLength + blockSize - 1) / blockSize) * blockSize;
        if (stream.ReadAtLeast(SegmentBytes, 0, encryptedLength, encryptedLength, throwOnEndOfStream: false) != encryptedLength)
            throw new EndOfStreamException("Incomplete encrypted package segment.");

        _ivBuffer ??= new byte[CryptoHelpers.MaxHashSize];
        var ivLength = Encryption.GenerateBlockKey(segmentIndex, IV, _ivBuffer);
        _decryptBuffer ??= new byte[SegmentLength];
#if NET8_0_OR_GREATER
        if (_cipher.Mode == System.Security.Cryptography.CipherMode.CBC)
        {
            // One-shot decryption reuses the cipher, IV and buffers instead of creating a decryptor and CryptoStream per segment.
            _cipher.DecryptCbc(SegmentBytes.AsSpan(0, encryptedLength), _ivBuffer.AsSpan(0, ivLength), _decryptBuffer, System.Security.Cryptography.PaddingMode.None);
        }
        else
#endif
        {
            Array.Resize(ref _ivBuffer, ivLength);
            using var decryptor = _cipher.CreateDecryptor(Key, _ivBuffer);
            if (decryptor.TransformBlock(SegmentBytes, 0, encryptedLength, _decryptBuffer, 0) != encryptedLength)
                throw new System.Security.Cryptography.CryptographicException("Incomplete encrypted package segment.");
        }

        (SegmentBytes, _decryptBuffer) = (_decryptBuffer, SegmentBytes);
        _loadedSegment = segmentIndex;
    }
}