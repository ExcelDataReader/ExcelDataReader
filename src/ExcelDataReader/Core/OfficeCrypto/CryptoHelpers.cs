using System;
using System.IO;
using System.Security.Cryptography;

namespace ExcelDataReader.Core.OfficeCrypto
{
    internal static class CryptoHelpers
    {
        public static HashAlgorithm Create(HashIdentifier hashAlgorithm) 
        {
            switch (hashAlgorithm)
            {
                case HashIdentifier.SHA512:
                    return SHA512.Create();
                case HashIdentifier.SHA384:
                    return SHA384.Create();
                case HashIdentifier.SHA256:
                    return SHA256.Create();
                case HashIdentifier.SHA1:
#pragma warning disable CA5350 // Do Not Use Weak Cryptographic Algorithms
                    return SHA1.Create();
#pragma warning restore CA5350 // Do Not Use Weak Cryptographic Algorithms
                case HashIdentifier.MD5:
#pragma warning disable CA5351 // Do Not Use Broken Cryptographic Algorithms
                    return MD5.Create();
#pragma warning restore CA5351 // Do Not Use Broken Cryptographic Algorithms
                default:
                    throw new InvalidOperationException("Unsupported hash algorithm");
            }
        }

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

        public static SymmetricAlgorithm CreateCipher(CipherIdentifier identifier, int keySize, int blockSize, CipherMode mode)
        {
            switch (identifier)
            {
                case CipherIdentifier.RC4:
                    return new RC4Managed();
                case CipherIdentifier.DES3:
#pragma warning disable CA5350 // Do Not Use Weak Cryptographic Algorithms
                    return InitCipher(TripleDES.Create(), keySize, blockSize, mode);
#pragma warning restore CA5350 // Do Not Use Weak Cryptographic Algorithms
#pragma warning disable CA5351 // Do Not Use Broken Cryptographic Algorithms
                case CipherIdentifier.RC2:
                    return InitCipher(RC2.Create(), keySize, blockSize, mode);
                case CipherIdentifier.DES:
                    return InitCipher(DES.Create(), keySize, blockSize, mode);
#pragma warning restore CA5351 // Do Not Use Broken Cryptographic Algorithms
                case CipherIdentifier.AES:
                    return InitCipher(new RijndaelManaged(), keySize, blockSize, mode);
            }

            throw new InvalidOperationException("Unsupported encryption method: " + identifier.ToString());
        }

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
}
