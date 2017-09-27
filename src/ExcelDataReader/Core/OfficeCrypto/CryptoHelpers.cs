using System;
using System.IO;
using System.Security.Cryptography;

namespace ExcelDataReader.Core.OfficeCrypto
{
    internal static class CryptoHelpers
    {
        public static byte[] HashBytes(byte[] bytes, HashIdentifier hashAlgorithm)
        {
            HashAlgorithm hash;
            if (hashAlgorithm == HashIdentifier.SHA512)
                hash = SHA512.Create();
            else if (hashAlgorithm == HashIdentifier.SHA384)
                hash = SHA384.Create();
            else if (hashAlgorithm == HashIdentifier.SHA256)
                hash = SHA256.Create();
            else if (hashAlgorithm == HashIdentifier.SHA1)
                hash = SHA1.Create();
            else if (hashAlgorithm == HashIdentifier.MD5)
                hash = MD5.Create();
            else
                throw new InvalidOperationException("Unsupported hash algorithm");

            using (hash)
            {
                return hash.ComputeHash(bytes);
            }
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
                    return InitCipher(TripleDES.Create(), keySize, blockSize, mode);
#if NET20 || NET45 || NETSTANDARD2_0
                case CipherIdentifier.RC2:
                    return InitCipher(RC2.Create(), keySize, blockSize, mode);
                case CipherIdentifier.DES:
                    return InitCipher(DES.Create(), keySize, blockSize, mode);
                case CipherIdentifier.AES:
                    return InitCipher(new RijndaelManaged(), keySize, blockSize, mode);
#else
                case CipherIdentifier.AES:
                    return InitCipher(Aes.Create(), keySize, blockSize, mode);
#endif
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
            using (var decryptor = algo.CreateDecryptor(key, iv))
            {
                return DecryptBytes(decryptor, bytes);
            }
        }

        public static byte[] DecryptBytes(ICryptoTransform transform, byte[] bytes)
        {
            var length = bytes.Length;
            using (MemoryStream msDecrypt = new MemoryStream(bytes, 0, length))
            {
                using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, transform, CryptoStreamMode.Read))
                {
                    var result = new byte[length];
                    csDecrypt.Read(result, 0, length);
                    return result;
                }
            }
        }
    }
}
