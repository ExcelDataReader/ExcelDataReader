using System;
using System.IO;
using System.Security.Cryptography;

namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Base class for the various encryption schemes used by Excel
    /// </summary>
    internal abstract class EncryptionInfo
    {
        /// <summary>
        /// Gets a value indicating whether XOR obfuscation is used.
        /// When true, the ICryptoTransform can be cast to XorTransform and
        /// handle the special case where XorArrayIndex must be manipulated
        /// per record.
        /// </summary>
        public abstract bool IsXor { get; }

        public static EncryptionInfo Create(ushort xorEncryptionKey, ushort xorHashValue)
        {
            return new XorEncryption()
            {
                EncryptionKey = xorEncryptionKey,
                HashValue = xorHashValue
            };
        }

        public static EncryptionInfo Create(byte[] bytes)
        {
            // TODO Does this work on a big endian system?
            var versionMajor = BitConverter.ToUInt16(bytes, 0);
            var versionMinor = BitConverter.ToUInt16(bytes, 2);

            if (versionMajor == 1 && versionMinor == 1)
            {
                return new RC4Encryption(bytes);
            }
            else if ((versionMajor == 2 || versionMajor == 3 || versionMajor == 4) && versionMinor == 2)
            {
                // 2.3.4.5 \EncryptionInfo Stream (Standard Encryption)
                return new StandardEncryption(bytes);
            }
            else if ((versionMajor == 3 || versionMajor == 4) && versionMinor == 3)
            {
                // 2.3.4.6 \EncryptionInfo Stream (Extensible Encryption)
                throw new InvalidOperationException("Extensible Encryption not supported");
            }
            else if (versionMajor == 4 && versionMinor == 4)
            {
                // 2.3.4.10 \EncryptionInfo Stream (Agile Encryption)
                return new AgileEncryption(bytes);
            }
            else
            {
                throw new InvalidOperationException("Unsupported EncryptionInfo version " + versionMajor + "." + versionMinor);
            }
        }

        public abstract byte[] GenerateSecretKey(string password);

        public abstract byte[] GenerateBlockKey(int blockNumber, byte[] secretKey);

        public abstract Stream CreateEncryptedPackageStream(Stream stream, byte[] secretKey);

        public abstract bool VerifyPassword(string password);

        public abstract SymmetricAlgorithm CreateCipher();
    }
}
