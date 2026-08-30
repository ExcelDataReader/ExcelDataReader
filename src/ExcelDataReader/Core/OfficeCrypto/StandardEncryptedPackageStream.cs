using System.Security.Cryptography;

namespace ExcelDataReader.Core.OfficeCrypto;

internal sealed class StandardEncryptedPackageStream : Stream
{
    public StandardEncryptedPackageStream(Stream underlyingStream, byte[] secretKey, StandardEncryption encryption)
    {
        Cipher = CryptoHelpers.CreateCipher(encryption.CipherAlgorithm, encryption.KeySize, encryption.BlockSize, CipherMode.ECB);
        Decryptor = Cipher.CreateDecryptor(secretKey, encryption.SaltValue);

        var header = new byte[8];
        underlyingStream.ReadAtLeast(header, 0, 8);
        DecryptedLength = BitConverter.ToInt32(header, 0);

        // Wrap CryptoStream to override the length and dispose the cipher and transform 
        // Zip readers scan backwards from the end for the central zip directory, and could fail if its too far away
        // CryptoStream is forward-only, so assume the zip readers read everything to memory
        BaseStream = new CryptoStream(underlyingStream, Decryptor, CryptoStreamMode.Read);
    }

    public override bool CanRead => GetBaseStream().CanRead;

    public override bool CanSeek => GetBaseStream().CanSeek;

    public override bool CanWrite => GetBaseStream().CanWrite;

    public override long Length => DecryptedLength;

    public override long Position
    {
        get => GetBaseStream().Position;
        set => GetBaseStream().Position = value;
    }

    private CryptoStream? BaseStream { get; set; }

    private SymmetricAlgorithm? Cipher { get; set; }

    private ICryptoTransform? Decryptor { get; set; }

    private long DecryptedLength { get; }

    public override void Flush()
    {
        GetBaseStream().Flush();
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        return GetBaseStream().Read(buffer, offset, count);
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        return GetBaseStream().Seek(offset, origin);
    }

    public override void SetLength(long value)
    {
        GetBaseStream().SetLength(value);
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        GetBaseStream().Write(buffer, offset, count);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            Decryptor?.Dispose();
            Decryptor = null;

            Cipher?.Dispose();
            Cipher = null;

            BaseStream?.Dispose();
            BaseStream = null;
        }

        base.Dispose(disposing);
    }

    private CryptoStream GetBaseStream()
    {
        return BaseStream ?? throw new ObjectDisposedException(nameof(StandardEncryptedPackageStream));
    }
}
