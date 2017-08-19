using System;
using System.IO;
using System.Security.Cryptography;
using System.Xml;

namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Represents "Agile Encryption" used in XLSX (Office 2010 and newer)
    /// </summary>
    internal class AgileEncryption : EncryptionInfo
    {
        private const string NEncryption = "encryption";
        private const string NKeyData = "keyData";
        private const string NKeyEncryptors = "keyEncryptors";
        private const string NKeyEncryptor = "keyEncryptor";
        private const string NEncryptedKey = "encryptedKey";
        private const string NsEncryption = "http://schemas.microsoft.com/office/2006/encryption";
        private const string NsPassword = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

        public AgileEncryption(byte[] bytes)
        {
            using (var stream = new MemoryStream(bytes, 8, bytes.Length - 8))
            {
                using (var xmlReader = XmlReader.Create(stream))
                {
                    ReadXmlEncryptionInfoStream(xmlReader);
                }
            }
        }

        public CipherIdentifier CipherAlgorithm { get; set; }

        public CipherMode CipherChaining { get; set; }

        public HashIdentifier HashAlgorithm { get; set; }

        public int KeyBits { get; set; }

        public int BlockSize { get; set; }

        public int HashSize { get; set; }

        public byte[] SaltValue { get; set; }

        public byte[] PasswordSaltValue { get; set; }

        public CipherIdentifier PasswordCipherAlgorithm { get; set; }

        public CipherMode PasswordCipherChaining { get; set; }

        public HashIdentifier PasswordHashAlgorithm { get; set; }

        public byte[] PasswordEncryptedKeyValue { get; set; }

        public byte[] PasswordEncryptedVerifierHashInput { get; set; }

        public byte[] PasswordEncryptedVerifierHashValue { get; set; }

        public int PasswordSpinCount { get; set; }

        public int PasswordKeyBits { get; set; }

        public int PasswordBlockSize { get; set; }

        public override bool IsXor => false;

        public override SymmetricAlgorithm CreateCipher()
        {
            return CryptoHelpers.CreateCipher(CipherAlgorithm, KeyBits, BlockSize * 8, CipherChaining);
        }

        public override byte[] GenerateSecretKey(string password)
        {
            using (var cipher = CryptoHelpers.CreateCipher(PasswordCipherAlgorithm, PasswordKeyBits, PasswordBlockSize * 8, PasswordCipherChaining))
            {
                return GenerateSecretKey(password, PasswordSaltValue, PasswordHashAlgorithm, PasswordEncryptedKeyValue, PasswordSpinCount, PasswordKeyBits, cipher);
            }
        }

        public override byte[] GenerateBlockKey(int blockNumber, byte[] secretKey)
        {
            var salt = CryptoHelpers.HashBytes(CryptoHelpers.Combine(secretKey, BitConverter.GetBytes(blockNumber)), HashAlgorithm);
            Array.Resize(ref salt, BlockSize);
            return salt;
        }

        public override Stream CreateEncryptedPackageStream(Stream stream, byte[] secretKey)
        {
            return new AgileEncryptedPackageStream(stream, secretKey, SaltValue, this);
        }

        public override bool VerifyPassword(string password)
        {
            var secretKey = HashPassword(password, PasswordSaltValue, PasswordHashAlgorithm, PasswordSpinCount);

            var inputBlockKey = CryptoHelpers.HashBytes(
                CryptoHelpers.Combine(secretKey, new byte[] { 0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79 }),
                PasswordHashAlgorithm);
            Array.Resize(ref inputBlockKey, PasswordKeyBits / 8);

            var valueBlockKey = CryptoHelpers.HashBytes(
                CryptoHelpers.Combine(secretKey, new byte[] { 0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e }),
                PasswordHashAlgorithm);
            Array.Resize(ref valueBlockKey, PasswordKeyBits / 8);

            using (var cipher = CryptoHelpers.CreateCipher(PasswordCipherAlgorithm, PasswordKeyBits, PasswordBlockSize * 8, PasswordCipherChaining))
            {
                var decryptedVerifier = CryptoHelpers.DecryptBytes(cipher, PasswordEncryptedVerifierHashInput, inputBlockKey, PasswordSaltValue);
                var decryptedVerifierHash = CryptoHelpers.DecryptBytes(cipher, PasswordEncryptedVerifierHashValue, valueBlockKey, PasswordSaltValue);

                var verifierHash = CryptoHelpers.HashBytes(decryptedVerifier, PasswordHashAlgorithm);
                for (var i = 0; i < Math.Min(decryptedVerifierHash.Length, verifierHash.Length); ++i)
                {
                    if (decryptedVerifierHash[i] != verifierHash[i])
                        return false;
                }

                return true;
            }
        }

        private static byte[] GenerateSecretKey(string password, byte[] saltValue, HashIdentifier hashAlgorithm, byte[] encryptedKeyValue, int spinCount, int keyBits, SymmetricAlgorithm cipher)
        {
            var block3 = new byte[] { 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6 };

            var h = HashPassword(password, saltValue, hashAlgorithm, spinCount);

            h = CryptoHelpers.HashBytes(CryptoHelpers.Combine(h, block3), hashAlgorithm);

            // Truncate or pad with 0x36
            var hashSize = h.Length;
            Array.Resize(ref h, keyBits / 8);
            for (var i = hashSize; i < keyBits / 8; i++)
            {
                h[i] = 0x36;
            }

            // NOTE: the stored salt is padded to a multiple of the block size which affects AES-192
            var decryptedKeyValue = CryptoHelpers.DecryptBytes(cipher, encryptedKeyValue, h, saltValue);
            Array.Resize(ref decryptedKeyValue, keyBits / 8);
            return decryptedKeyValue;
        }

        private static byte[] HashPassword(string password, byte[] saltValue, HashIdentifier hashAlgorithm, int spinCount)
        {
            var h = CryptoHelpers.HashBytes(CryptoHelpers.Combine(saltValue, System.Text.Encoding.Unicode.GetBytes(password)), hashAlgorithm);

            for (var i = 0; i < spinCount; i++)
            {
                h = CryptoHelpers.HashBytes(CryptoHelpers.Combine(BitConverter.GetBytes(i), h), hashAlgorithm);
            }

            return h;
        }

        private HashIdentifier ParseHash(string value)
        {
            return (HashIdentifier)Enum.Parse(typeof(HashIdentifier), value);
        }

        private CipherIdentifier ParseCipher(string value, int blockBits)
        {
            if (value == "AES")
            {
                return CipherIdentifier.AES;
            }
            else if (value == "DES")
            {
                return CipherIdentifier.DES;
            }
            else if (value == "3DES")
            {
                return CipherIdentifier.DES3;
            }
            else if (value == "RC2")
            {
                return CipherIdentifier.RC2;
            }

            throw new ArgumentException(nameof(value), "Unknown encryption: " + value);
        }

        private CipherMode ParseCipherMode(string value)
        {
            if (value == "ChainingModeCBC")
                return CipherMode.CBC;
#if NET20 || NET45 || NETSTANDARD2_0
            else if (value == "ChainingModeCFB")
                return CipherMode.CFB;
#endif
            throw new ArgumentException("Invalid CipherMode " + value);
        }

        private void ReadXmlEncryptionInfoStream(XmlReader xmlReader)
        {
            if (!xmlReader.IsStartElement(NEncryption, NsEncryption))
            {
                return;
            }

            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                return;
            }

            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NKeyData, NsEncryption))
                {
                    // <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="zYmgeIEW4PVmYPiNJItVCQ=="/>
                    // <dataIntegrity encryptedHmacKey="v11xCwbBfQ6Wq03h2M6Nh5Z9fwNnFQwEzu8vmBDps55kd+HfLDzrnuKzuQq4tlpxW0nX99VWh+n2X6ukU6v9FQ==" encryptedHmacValue="SvDwFQR4dNsXOzNstFWHqSHpAUWHQvAr63IhxlxhlQEAczDPIwCWD32aIEFipY7NOlW+LvYPaKC8zO1otxit2g=="/>
                    // <keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="n37HW2mNfJuGwVxTeBY1LA==" encryptedVerifierHashInput="2Y2Oo+QDyMdo327gZUcejA==" encryptedVerifierHashValue="PmkCD5y5cHqMQqbgACUgxLRgISYZL6+jj3K0PSrFDWlEG+fjzFevIee1FubgdpY2P22IIM6W7C/bXE0ayAo8yg==" encryptedKeyValue="qzkvVPIBy2Bk/w2/fp+hhpq5sPReA8aUu414/Xh7494="/></keyEncryptor></keyEncryptors>
                    int saltSize, blockSize, keyBits, hashSize;
                    var cipherAlgorithm = xmlReader.GetAttribute("cipherAlgorithm");
                    var cipherChaining = xmlReader.GetAttribute("cipherChaining");
                    var hashAlgorithm = xmlReader.GetAttribute("hashAlgorithm");
                    var saltValue = xmlReader.GetAttribute("saltValue");

                    int.TryParse(xmlReader.GetAttribute("saltSize"), out saltSize);
                    int.TryParse(xmlReader.GetAttribute("blockSize"), out blockSize);
                    int.TryParse(xmlReader.GetAttribute("keyBits"), out keyBits);
                    int.TryParse(xmlReader.GetAttribute("hashSize"), out hashSize);

                    SaltValue = Convert.FromBase64String(saltValue);
                    HashSize = hashSize; // given in bytes, also given implicitly by SHA512
                    KeyBits = keyBits;
                    BlockSize = blockSize;
                    CipherAlgorithm = ParseCipher(cipherAlgorithm, blockSize * 8);
                    CipherChaining = ParseCipherMode(cipherChaining);
                    HashAlgorithm = ParseHash(hashAlgorithm);
                    xmlReader.Skip();
                }
                else if (xmlReader.IsStartElement(NKeyEncryptors, NsEncryption))
                {
                    ReadKeyEncryptors(xmlReader);
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }
        }

        private void ReadKeyEncryptors(XmlReader xmlReader)
        {
            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                return;
            }

            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NKeyEncryptor, NsEncryption))
                {
                    // <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                    ReadKeyEncryptor(xmlReader);
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }
        }

        private void ReadKeyEncryptor(XmlReader xmlReader)
        {
            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                return;
            }

            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NEncryptedKey, NsPassword))
                {
                    // <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="n37HW2mNfJuGwVxTeBY1LA==" encryptedVerifierHashInput="2Y2Oo+QDyMdo327gZUcejA==" encryptedVerifierHashValue="PmkCD5y5cHqMQqbgACUgxLRgISYZL6+jj3K0PSrFDWlEG+fjzFevIee1FubgdpY2P22IIM6W7C/bXE0ayAo8yg==" encryptedKeyValue="qzkvVPIBy2Bk/w2/fp+hhpq5sPReA8aUu414/Xh7494="/></keyEncryptor></keyEncryptors>
                    int spinCount, saltSize, blockSize, keyBits, hashSize;
                    var cipherAlgorithm = xmlReader.GetAttribute("cipherAlgorithm");
                    var cipherChaining = xmlReader.GetAttribute("cipherChaining");
                    var hashAlgorithm = xmlReader.GetAttribute("hashAlgorithm");
                    var saltValue = xmlReader.GetAttribute("saltValue");
                    var encryptedVerifierHashInput = xmlReader.GetAttribute("encryptedVerifierHashInput");
                    var encryptedVerifierHashValue = xmlReader.GetAttribute("encryptedVerifierHashValue");
                    var encryptedKeyValue = xmlReader.GetAttribute("encryptedKeyValue");

                    int.TryParse(xmlReader.GetAttribute("spinCount"), out spinCount);
                    int.TryParse(xmlReader.GetAttribute("saltSize"), out saltSize);
                    int.TryParse(xmlReader.GetAttribute("blockSize"), out blockSize);
                    int.TryParse(xmlReader.GetAttribute("keyBits"), out keyBits);
                    int.TryParse(xmlReader.GetAttribute("hashSize"), out hashSize);

                    PasswordSaltValue = Convert.FromBase64String(saltValue);
                    PasswordCipherAlgorithm = ParseCipher(cipherAlgorithm, blockSize * 8);
                    PasswordCipherChaining = ParseCipherMode(cipherChaining);
                    PasswordHashAlgorithm = ParseHash(hashAlgorithm);
                    PasswordEncryptedKeyValue = Convert.FromBase64String(encryptedKeyValue);
                    PasswordEncryptedVerifierHashInput = Convert.FromBase64String(encryptedVerifierHashInput);
                    PasswordEncryptedVerifierHashValue = Convert.FromBase64String(encryptedVerifierHashValue);
                    PasswordSpinCount = spinCount;
                    PasswordKeyBits = keyBits;
                    PasswordBlockSize = blockSize;

                    xmlReader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }
        }
    }
}
