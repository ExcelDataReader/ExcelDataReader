using System;
using System.IO;
using System.Security.Cryptography;

namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Represents "XOR Deobfucation Method 1" used in XLS.
    /// </summary>
    internal class XorEncryption : EncryptionInfo
    {
        public ushort EncryptionKey { get; set; }

        public ushort HashValue { get; set; }

        public override bool IsXor => true;

        public override SymmetricAlgorithm CreateCipher()
        {
            return new XorManaged();
        }

        public override Stream CreateEncryptedPackageStream(Stream stream, byte[] secretKey)
        {
            throw new NotImplementedException();
        }

        public override byte[] GenerateBlockKey(int blockNumber, byte[] secretKey)
        {
            return secretKey;
        }

        public override byte[] GenerateSecretKey(string password)
        {
            var passwordBytes = System.Text.Encoding.ASCII.GetBytes(password.Substring(0, Math.Min(password.Length, 15)));
            return XorManaged.CreateXorArray_Method1(passwordBytes);
        }

        public override bool VerifyPassword(string password)
        {
            var passwordBytes = System.Text.Encoding.ASCII.GetBytes(password.Substring(0, Math.Min(password.Length, 15)));
            var verifier = XorManaged.CreatePasswordVerifier_Method1(passwordBytes);
            return verifier == HashValue;
        }
    }
}
