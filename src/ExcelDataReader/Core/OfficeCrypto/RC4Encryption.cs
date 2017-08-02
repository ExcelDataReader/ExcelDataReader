using System;
using System.IO;
using System.Security.Cryptography;

namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Represents the binary RC4+MD5 encryption header used in XLS.
    /// </summary>
    internal class RC4Encryption : EncryptionInfo
    {
        public RC4Encryption(byte[] bytes)
        {
            Salt = new byte[16];
            EncryptedVerifier = new byte[16];
            EncryptedVerifierHash = new byte[16];
            Array.Copy(bytes, 4, Salt, 0, 16);
            Array.Copy(bytes, 4 + 16, EncryptedVerifier, 0, 16);
            Array.Copy(bytes, 4 + 32, EncryptedVerifierHash, 0, 16);
        }

        public byte[] Salt { get; }

        public byte[] EncryptedVerifier { get; }

        public byte[] EncryptedVerifierHash { get; }

        public override bool IsXor => false;

        public static byte[] GenerateSecretKey(string password, byte[] salt)
        {
            if (password.Length > 16)
                password = password.Substring(0, 16);
            var h = CryptoHelpers.HashBytes(System.Text.Encoding.Unicode.GetBytes(password), HashIdentifier.MD5);
            Array.Resize(ref h, 5);

            // Combine h + salt 16 times:
            h = CryptoHelpers.Combine(h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt, h, salt);
            h = CryptoHelpers.HashBytes(h, HashIdentifier.MD5);
            Array.Resize(ref h, 5);
            return h;
        }

        public override SymmetricAlgorithm CreateCipher()
        {
            return CryptoHelpers.CreateCipher(CipherIdentifier.RC4, 0, 0, 0);
        }

        public override Stream CreateEncryptedPackageStream(Stream stream, byte[] secretKey)
        {
            throw new NotImplementedException();
        }

        public override byte[] GenerateBlockKey(int blockNumber, byte[] secretKey)
        {
            var salt = CryptoHelpers.Combine(secretKey, BitConverter.GetBytes(blockNumber));
            return CryptoHelpers.HashBytes(salt, HashIdentifier.MD5);
        }

        public override byte[] GenerateSecretKey(string password)
        {
            return GenerateSecretKey(password, Salt);
        }

        public override bool VerifyPassword(string password)
        {
            // 2.3.6.4 Password Verification
            var secretKey = GenerateSecretKey(password);
            var blockKey = GenerateBlockKey(0, secretKey);

            using (var cipher = CryptoHelpers.CreateCipher(CipherIdentifier.RC4, 0, 0, 0))
            {
                using (var transform = cipher.CreateDecryptor(blockKey, null))
                {
                    var decryptedVerifier = CryptoHelpers.DecryptBytes(transform, EncryptedVerifier);
                    var decryptedVerifierHash = CryptoHelpers.DecryptBytes(transform, EncryptedVerifierHash);

                    var verifierHash = CryptoHelpers.HashBytes(decryptedVerifier, HashIdentifier.MD5);
                    for (var i = 0; i < 16; ++i)
                    {
                        if (decryptedVerifierHash[i] != verifierHash[i])
                            return false;
                    }

                    return true;
                }
            }
        }
    }
}
