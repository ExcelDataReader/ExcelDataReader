using System.Security.Cryptography;

namespace ExcelDataReader.Core.OfficeCrypto;

internal static class CryptoHelpers
{
    /// <summary>
    /// Largest supported hash (SHA-512) in bytes; also an upper bound for block key lengths.
    /// </summary>
    public const int MaxHashSize = 64;

    public static HashAlgorithm Create(HashIdentifier hashAlgorithm) => hashAlgorithm switch
    {
        HashIdentifier.SHA512 => SHA512.Create(),
        HashIdentifier.SHA384 => SHA384.Create(),
        HashIdentifier.SHA256 => SHA256.Create(),
#pragma warning disable CA5350 // Do Not Use Weak Cryptographic Algorithms
        HashIdentifier.SHA1 => SHA1.Create(),
#pragma warning restore CA5350 // Do Not Use Weak Cryptographic Algorithms
#pragma warning disable CA5351 // Do Not Use Broken Cryptographic Algorithms
        HashIdentifier.MD5 => MD5.Create(),
#pragma warning restore CA5351 // Do Not Use Broken Cryptographic Algorithms
        _ => throw new InvalidOperationException("Unsupported hash algorithm"),
    };

    public static byte[] HashBytes(byte[] bytes, HashIdentifier hashAlgorithm)
    {
        using HashAlgorithm hash = Create(hashAlgorithm);
        return hash.ComputeHash(bytes);
    }

    public static byte[] Combine(byte[] first, byte[] second) => [.. first, .. second];

    public static void WriteInt32LittleEndian(byte[] destination, int offset, int value)
    {
        destination[offset] = (byte)value;
        destination[offset + 1] = (byte)(value >> 8);
        destination[offset + 2] = (byte)(value >> 16);
        destination[offset + 3] = (byte)(value >> 24);
    }

    /// <summary>
    /// Writes hash(prefix || LE32(blockNumber)) into <paramref name="destination"/>, truncated to
    /// <paramref name="keyLength"/> bytes and zero-padded to <paramref name="resultLength"/> bytes.
    /// <paramref name="keyLength"/> must not exceed <paramref name="resultLength"/>.
    /// </summary>
    public static void HashBlockKey(byte[] prefix, int blockNumber, HashIdentifier hashAlgorithm, int keyLength, byte[] destination, int resultLength)
    {
#if NET8_0_OR_GREATER
        Span<byte> input = stackalloc byte[prefix.Length + 4];
        prefix.CopyTo(input);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(input[prefix.Length..], blockNumber);

        Span<byte> hash = stackalloc byte[MaxHashSize];
        var hashLength = hashAlgorithm switch
        {
            HashIdentifier.SHA512 => SHA512.HashData(input, hash),
            HashIdentifier.SHA384 => SHA384.HashData(input, hash),
            HashIdentifier.SHA256 => SHA256.HashData(input, hash),
#pragma warning disable CA5350 // Do Not Use Weak Cryptographic Algorithms
            HashIdentifier.SHA1 => SHA1.HashData(input, hash),
#pragma warning restore CA5350 // Do Not Use Weak Cryptographic Algorithms
#pragma warning disable CA5351 // Do Not Use Broken Cryptographic Algorithms
            HashIdentifier.MD5 => MD5.HashData(input, hash),
#pragma warning restore CA5351 // Do Not Use Broken Cryptographic Algorithms
            _ => throw new InvalidOperationException("Unsupported hash algorithm"),
        };

        var copyLength = Math.Min(hashLength, keyLength);
        hash[..copyLength].CopyTo(destination);
#else
        var input = new byte[prefix.Length + 4];
        Buffer.BlockCopy(prefix, 0, input, 0, prefix.Length);
        WriteInt32LittleEndian(input, prefix.Length, blockNumber);

        byte[] hash;
        using (var algorithm = Create(hashAlgorithm))
            hash = algorithm.ComputeHash(input);

        var copyLength = Math.Min(hash.Length, keyLength);
        Buffer.BlockCopy(hash, 0, destination, 0, copyLength);
#endif
        Array.Clear(destination, copyLength, resultLength - copyLength);
    }

    /// <summary>
    /// Runs the key derivation spin loop <c>hash = H(LE32(i) || hash)</c> in place.
    /// <paramref name="hash"/> must be exactly one hash long.
    /// </summary>
    public static void SpinHash(HashAlgorithm hashAlgorithm, byte[] hash, int spinCount)
    {
#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
        Span<byte> iterationData = stackalloc byte[4 + hash.Length];
        for (var i = 0; i < spinCount; i++)
        {
            System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(iterationData[..4], i);
            hash.CopyTo(iterationData[4..]);
            hashAlgorithm.TryComputeHash(iterationData, hash, out _);
        }
#else
        var iterationData = new byte[4 + hash.Length];
        for (var i = 0; i < spinCount; i++)
        {
            WriteInt32LittleEndian(iterationData, 0, i);
            Buffer.BlockCopy(hash, 0, iterationData, 4, hash.Length);
            Buffer.BlockCopy(hashAlgorithm.ComputeHash(iterationData), 0, hash, 0, hash.Length);
        }
#endif
    }

    public static SymmetricAlgorithm CreateCipher(CipherIdentifier identifier, int keySize, int blockSize, CipherMode mode) => identifier switch 
    {
        CipherIdentifier.RC4 => new RC4Managed(),
#pragma warning disable CA5350 // Do Not Use Weak Cryptographic Algorithms
        CipherIdentifier.DES3 => InitCipher(TripleDES.Create(), keySize, blockSize, mode),
#pragma warning restore CA5350 // Do Not Use Weak Cryptographic Algorithms
#pragma warning disable CA5351 // Do Not Use Broken Cryptographic Algorithms
        CipherIdentifier.RC2 => InitCipher(RC2.Create(), keySize, blockSize, mode),
        CipherIdentifier.DES => InitCipher(DES.Create(), keySize, blockSize, mode),
#pragma warning restore CA5351 // Do Not Use Broken Cryptographic Algorithms
        CipherIdentifier.AES => InitCipher(Aes.Create(), keySize, blockSize, mode),
        _ => throw new InvalidOperationException("Unsupported encryption method: " + identifier.ToString()),
    };

    public static SymmetricAlgorithm InitCipher(SymmetricAlgorithm cipher, int keySize, int blockSize, CipherMode mode)
    {
        cipher.KeySize = keySize;
        cipher.BlockSize = blockSize;
        cipher.Mode = mode;
        cipher.Padding = PaddingMode.Zeros;
        return cipher;
    }

    public static byte[] DecryptBytes(SymmetricAlgorithm algo, byte[] bytes, byte[] key, byte[] iv)
    {
        using var decryptor = algo.CreateDecryptor(key, iv);
        return DecryptBytes(decryptor, bytes, bytes.Length);
    }

    public static byte[] DecryptBytes(ICryptoTransform transform, byte[] bytes, int chunkSize)
    {
        var length = chunkSize;
        using MemoryStream msDecrypt = new(bytes, 0, length);
        using CryptoStream csDecrypt = new(msDecrypt, transform, CryptoStreamMode.Read);
        var result = new byte[length];
        csDecrypt.ReadAtLeast(result, 0, length);
        return result;
    }

    public static void DecryptBytes(ICryptoTransform transform, byte[] bytes, int inputOffset, int chunkSize, byte[] output)
    {
        using MemoryStream msDecrypt = new(bytes, inputOffset, chunkSize);
        using CryptoStream csDecrypt = new(msDecrypt, transform, CryptoStreamMode.Read);
        csDecrypt.ReadAtLeast(output, 0, chunkSize);
    }
}
