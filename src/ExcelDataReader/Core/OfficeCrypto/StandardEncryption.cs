using System;
using System.IO;
using System.Security.Cryptography;

namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Represents the binary "Standard Encryption" header used in XLS and XLSX.
    /// XLS uses RC4+SHA1. XLSX uses AES+SHA1.
    /// </summary>
    internal class StandardEncryption : EncryptionInfo
    {
        private const int AesBlockSize = 128;
        private const int RC4BlockSize = 8;

        public StandardEncryption(byte[] bytes)
        {
            Flags = (EncryptionHeaderFlags)BitConverter.ToUInt32(bytes, 4);

            var headerSize = BitConverter.ToInt32(bytes, 8);

            // Using ProviderType and KeySize instead
            var cipher = (StandardCipher)BitConverter.ToUInt32(bytes, 20);

            var hashAlgorithm = (StandardHash)BitConverter.ToUInt32(bytes, 24);

            if ((Flags & EncryptionHeaderFlags.External) == 0)
            {
                switch (hashAlgorithm)
                {
                    case StandardHash.Default:
                    case StandardHash.SHA1:
                        HashAlgorithm = HashIdentifier.SHA1;
                        break;
                }
            }

            // ECMA-376: 0x00000080 (AES-128), 0x000000C0 (AES-192), or 0x00000100 (AES-256).
            // RC4: 0x00000028 – 0x00000080 (inclusive), 8-bits increments
            KeySize = BitConverter.ToInt32(bytes, 28);

            // Don't use this; is implementation-specific
            var providerType = (StandardProvider)BitConverter.ToUInt32(bytes, 32);

            // skip two reserved dwords
            CSPName = System.Text.Encoding.Unicode.GetString(bytes, 44, headerSize - 44 + 12); // +12 because we start counting from the offset after HeaderSize

            var saltSize = BitConverter.ToInt32(bytes, 12 + headerSize);

            SaltValue = new byte[saltSize];
            Array.Copy(bytes, 12 + headerSize + 4, SaltValue, 0, saltSize);

            Verifier = new byte[16];
            Array.Copy(bytes, 12 + headerSize + 4 + saltSize, Verifier, 0, 16);

            // An unsigned integer that specifies the number of bytes needed to
            // contain the hash of the data used to generate the EncryptedVerifier field.
            VerifierHashBytesNeeded = BitConverter.ToInt32(bytes, 12 + headerSize + 4 + saltSize + 16);

            // If the encryption algorithm is RC4, the length MUST be 20 bytes. If the encryption algorithm is AES, the length MUST be 32 bytes
            var verifierHashSize = ((Flags & EncryptionHeaderFlags.AES) != 0) ? 32 : 20;

            if (cipher == StandardCipher.RC4)
            {
                BlockSize = RC4BlockSize;
                verifierHashSize = 20;
            }
            else if (cipher == StandardCipher.AES128 || cipher == StandardCipher.AES192 || cipher == StandardCipher.AES256)
            {
                BlockSize = AesBlockSize;
                verifierHashSize = 32;
            }

            VerifierHash = new byte[verifierHashSize];
            Array.Copy(bytes, 12 + headerSize + 4 + saltSize + 16 + 4, VerifierHash, 0, verifierHashSize);

            if ((Flags & EncryptionHeaderFlags.External) == 0)
            {
                switch (cipher)
                {
                    case StandardCipher.Default:
                        if ((Flags & EncryptionHeaderFlags.AES) != 0)
                        {
                            CipherAlgorithm = CipherIdentifier.AES;
                        }
                        else if ((Flags & EncryptionHeaderFlags.CryptoAPI) != 0)
                        {
                            CipherAlgorithm = CipherIdentifier.RC4;
                        }

                        break;
                    case StandardCipher.AES128:
                    case StandardCipher.AES192:
                    case StandardCipher.AES256:
                        CipherAlgorithm = CipherIdentifier.AES;
                        break;

                    case StandardCipher.RC4:
                        CipherAlgorithm = CipherIdentifier.RC4;
                        break;
                }
            }
        }

        private enum StandardProvider
        {
            Default = 0x00000000,
            RC4 = 0x00000001,
            AES = 0x00000018,
        }

        private enum StandardCipher
        {
            Default = 0x00000000,
            AES128 = 0x0000660E,
            AES192 = 0x0000660F,
            AES256 = 0x00006610,
            RC4 = 0x00006801
        }

        private enum StandardHash
        {
            Default = 0x00000000,
            SHA1 = 0x00008004,
        }

        private enum EncryptionHeaderFlags : uint
        {
            CryptoAPI = 0x00000004,
            DocProps = 0x00000008,
            External = 0x00000010,
            AES = 0x00000020,
        }

        public CipherIdentifier CipherAlgorithm { get; set; }

        public HashIdentifier HashAlgorithm { get; set; }

        public int BlockSize { get; set; }

        public int KeySize { get; set; }

        public string CSPName { get; set; }

        public byte[] SaltValue { get; set; }

        public byte[] Verifier { get; set; }

        public byte[] VerifierHash { get; set; }

        public int VerifierHashBytesNeeded { get; set; }

        public override bool IsXor => false;

        private EncryptionHeaderFlags Flags { get; set; }

        public override SymmetricAlgorithm CreateCipher()
        {
            return CryptoHelpers.CreateCipher(CipherAlgorithm, KeySize, BlockSize, CipherMode.ECB);
        }

        public override Stream CreateEncryptedPackageStream(Stream stream, byte[] secretKey)
        {
            return new StandardEncryptedPackageStream(stream, secretKey, this);
        }

        public override byte[] GenerateBlockKey(int blockNumber, byte[] secretKey)
        {
            if ((Flags & EncryptionHeaderFlags.AES) != 0)
            {
                /*var salt = CryptoHelpers.Combine(secretKey, BitConverter.GetBytes(blockNumber));
                salt = CryptoHelpers.HashBytes(salt, HashAlgorithm);
                Array.Resize(ref salt, (int)KeySize / 8);
                return salt;*/
                throw new Exception("Block key for ECMA-376 Standard Encryption not implemented");
            }
            else if ((Flags & EncryptionHeaderFlags.CryptoAPI) != 0)
            {
                var salt = CryptoHelpers.Combine(secretKey, BitConverter.GetBytes(blockNumber));
                salt = CryptoHelpers.HashBytes(salt, HashAlgorithm);
                Array.Resize(ref salt, (int)KeySize / 8);
                return salt;
            }
            else
            {
                throw new InvalidOperationException("Unknown encryption type");
            }
        }

        public override byte[] GenerateSecretKey(string password)
        {
            if ((Flags & EncryptionHeaderFlags.AES) != 0)
            {
                // 2.3.4.7 ECMA-376 Document Encryption Key Generation (Standard Encryption)
                return GenerateEcma376SecretKey(password, SaltValue, HashAlgorithm, (int)KeySize, VerifierHashBytesNeeded);
            }
            else if ((Flags & EncryptionHeaderFlags.CryptoAPI) != 0)
            {
                // 2.3.5.2 RC4 CryptoAPI Encryption Key Generation
                return GenerateCryptoApiSecretKey(password, SaltValue, HashAlgorithm, (int)KeySize);
            }
            else
            {
                throw new InvalidOperationException("Unknown encryption type");
            }
        }

        public override bool VerifyPassword(string password)
        {
            // 2.3.4.9 Password Verification (Standard Encryption)
            // 2.3.5.6 Password Verification
            var secretKey = GenerateSecretKey(password);

            var blockKey = ((Flags & EncryptionHeaderFlags.AES) != 0) ? secretKey : GenerateBlockKey(0, secretKey);

            using (var cipher = CryptoHelpers.CreateCipher(CipherAlgorithm, KeySize, BlockSize, CipherMode.ECB))
            {
                using (var transform = cipher.CreateDecryptor(blockKey, SaltValue))
                {
                    var decryptedVerifier = CryptoHelpers.DecryptBytes(transform, Verifier);
                    var decryptedVerifierHash = CryptoHelpers.DecryptBytes(transform, VerifierHash);

                    var verifierHash = CryptoHelpers.HashBytes(decryptedVerifier, HashAlgorithm);
                    for (var i = 0; i < 16; ++i)
                    {
                        if (decryptedVerifierHash[i] != verifierHash[i])
                            return false;
                    }

                    return true;
                }
            }
        }

        /// <summary>
        /// 2.3.5.2 RC4 CryptoAPI Encryption Key Generation
        /// </summary>
        private static byte[] GenerateCryptoApiSecretKey(string password, byte[] saltValue, HashIdentifier hashAlgorithm, int keySize)
        {
            return CryptoHelpers.HashBytes(CryptoHelpers.Combine(saltValue, System.Text.Encoding.Unicode.GetBytes(password)), hashAlgorithm);
        }

        /// <summary>
        /// 2.3.4.7 ECMA-376 Document Encryption Key Generation (Standard Encryption)
        /// </summary>
        private static byte[] GenerateEcma376SecretKey(string password, byte[] saltValue, HashIdentifier hashAlgorithm, int keySize, int verifierHashSize)
        {
            var h = CryptoHelpers.HashBytes(CryptoHelpers.Combine(saltValue, System.Text.Encoding.Unicode.GetBytes(password)), hashAlgorithm);
            for (int i = 0; i < 50000; i++)
            {
                h = CryptoHelpers.HashBytes(CryptoHelpers.Combine(BitConverter.GetBytes(i), h), hashAlgorithm);
            }

            h = CryptoHelpers.HashBytes(CryptoHelpers.Combine(h, BitConverter.GetBytes(0)), hashAlgorithm);

            // The algorithm in this 'DeriveKey' function is the bit that's not clear from the documentation
            h = DeriveKey(h, hashAlgorithm, keySize, verifierHashSize);

            Array.Resize(ref h, keySize / 8);

            return h;
        }

        private static byte[] DeriveKey(byte[] hashValue, HashIdentifier hashAlgorithm, int keySize, int verifierHashSize)
        {
            // And one more hash to derive the key
            byte[] derivedKey = new byte[64];

            // This is step 4a in 2.3.4.7 of MS_OFFCRYPT version 1.0
            // and is required even though the notes say it should be 
            // used only when the encryption algorithm key > hash length.
            for (int i = 0; i < derivedKey.Length; i++)
                derivedKey[i] = (byte)(i < hashValue.Length ? 0x36 ^ hashValue[i] : 0x36);

            byte[] x1 = CryptoHelpers.HashBytes(derivedKey, hashAlgorithm);

            if (verifierHashSize > keySize / 8)
                return x1;

            for (int i = 0; i < derivedKey.Length; i++)
                derivedKey[i] = (byte)(i < hashValue.Length ? 0x5C ^ hashValue[i] : 0x5C);

            byte[] x2 = CryptoHelpers.HashBytes(derivedKey, hashAlgorithm);
            return CryptoHelpers.Combine(x1, x2);
        }
    }
}