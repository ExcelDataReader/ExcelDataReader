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

    public AgileEncryptedPackageStream(Stream stream, byte[] key, byte[] iv, EncryptionInfo encryption)
    {
        Stream = stream;
        Key = key;
        IV = iv;
        Encryption = encryption;

        Stream.ReadAtLeast(SegmentBytes, 0, 8);
        DecryptedLength = BitConverter.ToInt32(SegmentBytes, 0);
        ReadSegment();
    }

    public override bool CanRead => true;

    public override bool CanSeek => true;

    public override bool CanWrite => false;

    public override long Length => DecryptedLength;

    public override long Position { get => Offset - SegmentLength + SegmentOffset; set => Seek(value, SeekOrigin.Begin); }

    private Stream? Stream { get; set; }

    private byte[] Key { get; }

    private byte[] IV { get; }

    private EncryptionInfo Encryption { get; }

    private int Offset { get; set; }

    private byte[] SegmentBytes { get; set; } = new byte[SegmentLength];

    private int SegmentOffset { get; set; }

    private int SegmentIndex { get; set; }

    private int DecryptedLength { get; set; }

    public override void Flush()
    {
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        return Read(buffer.AsSpan(offset, count));
#else
        if (Position >= Length)
        {
            throw new InvalidOperationException("Tried to read past the end of the encrypted stream");
        }

        int index = 0;
        while (index < count)
        {
            if (SegmentOffset == SegmentBytes.Length)
            {
                ReadSegment();
                SegmentOffset = 0;
            }

            var chunkSize = Math.Min(count - index, SegmentBytes.Length - SegmentOffset);
            Array.Copy(SegmentBytes, SegmentOffset, buffer, offset + index, chunkSize);
            index += chunkSize;
            SegmentOffset += chunkSize;
        }

        return index;
#endif
    }

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
    public override int Read(Span<byte> buffer)
    {
        if (Position >= Length)
        {
            throw new InvalidOperationException("Tried to read past the end of the encrypted stream");
        }

        int index = 0;
        while (index < buffer.Length)
        {
            if (SegmentOffset == SegmentBytes.Length)
            {
                ReadSegment();
                SegmentOffset = 0;
            }

            var chunkSize = Math.Min(buffer.Length - index, SegmentBytes.Length - SegmentOffset);
            SegmentBytes.AsSpan(SegmentOffset, chunkSize).CopyTo(buffer[index..]);
            index += chunkSize;
            SegmentOffset += chunkSize;
        }

        return index;
    }
#endif

    public override long Seek(long offset, SeekOrigin origin)
    {
        switch (origin)
        {
            case SeekOrigin.Begin:
                SegmentIndex = (int)(offset / SegmentLength);
                Offset = SegmentIndex * SegmentLength;
                SegmentOffset = (int)(offset % SegmentLength);
                if (Offset < Length)
                    ReadSegment();
                return Position;
            case SeekOrigin.Current:
                return Seek(Position + offset, SeekOrigin.Begin);
            case SeekOrigin.End:
                return Seek(Length + offset, SeekOrigin.Begin);
            default:
                return Offset;
        }
    }

    public override void SetLength(long value)
    {
        throw new NotImplementedException();
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        throw new NotImplementedException();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            Stream?.Dispose();
            Stream = null;
            _cipher?.Dispose();
            _cipher = null;
        }

        base.Dispose(disposing);
    }

    private void ReadSegment()
    {
        var stream = Stream ?? throw new ObjectDisposedException(nameof(AgileEncryptedPackageStream));

        // NOTE: +8 skips EncryptedPackage header
        stream.Seek(8 + Offset, SeekOrigin.Begin);
        stream.ReadAtLeast(SegmentBytes, 0, SegmentLength);

        if (_cipher == null)
        {
            _cipher = Encryption.CreateCipher();
            _cipher.Key = Key;
            _cipher.Padding = System.Security.Cryptography.PaddingMode.None;
        }

        _ivBuffer ??= new byte[CryptoHelpers.MaxHashSize];
        var ivLength = Encryption.GenerateBlockKey(SegmentIndex, IV, _ivBuffer);
        _decryptBuffer ??= new byte[SegmentLength];
#if NET8_0_OR_GREATER
        if (_cipher.Mode == System.Security.Cryptography.CipherMode.CBC)
        {
            // One-shot decryption reuses the cipher, IV and buffers instead of creating a decryptor and CryptoStream per segment.
            _cipher.DecryptCbc(SegmentBytes, _ivBuffer.AsSpan(0, ivLength), _decryptBuffer, System.Security.Cryptography.PaddingMode.None);
        }
        else
#endif
        {
            Array.Resize(ref _ivBuffer, ivLength);
            using var decryptor = _cipher.CreateDecryptor(Key, _ivBuffer);
            if (decryptor.TransformBlock(SegmentBytes, 0, SegmentLength, _decryptBuffer, 0) != SegmentLength)
                throw new System.Security.Cryptography.CryptographicException("Incomplete encrypted package segment.");
        }

        (SegmentBytes, _decryptBuffer) = (_decryptBuffer, SegmentBytes);
        SegmentIndex++;
        Offset += SegmentLength;
    }
}