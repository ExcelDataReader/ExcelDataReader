using System.Security.Cryptography;

namespace ExcelDataReader.Core.OfficeCrypto;

internal static class CryptoHelpers
{
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

    public static byte[] Combine(params byte[][] arrays)
    {
        var length = 0;
        for (var i = 0; i < arrays.Length; i++)
            length += arrays[i].Length;

        byte[] ret = new byte[length];
        int offset = 0;
        foreach (byte[] data in arrays)
        {
            Buffer.BlockCopy(data, 0, ret, offset, data.Length);
            offset += data.Length;
        }

        return ret;
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
        return DecryptBytes(decryptor, bytes);
    }

    public static byte[] DecryptBytes(ICryptoTransform transform, byte[] bytes)
    {
        var length = bytes.Length;
        using MemoryStream msDecrypt = new(bytes, 0, length);
        using CryptoStream csDecrypt = new(msDecrypt, transform, CryptoStreamMode.Read);
        var result = new byte[length];
        csDecrypt.ReadAtLeast(result, 0, length);
        return result;
    }
}
