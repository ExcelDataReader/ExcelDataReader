using System.Security.Cryptography;

namespace ExcelDataReader.Core.OfficeCrypto;

/// <summary>
/// Represents the binary RC4+MD5 encryption header used in XLS.
/// </summary>
internal sealed class RC4Encryption : EncryptionInfo
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

        // 2.3.6.2: concatenate h and the salt 16 times, then hash and truncate to 5 bytes.
        var combined = new byte[16 * (h.Length + salt.Length)];
        for (var i = 0; i < 16; i++)
        {
            var offset = i * (h.Length + salt.Length);
            Buffer.BlockCopy(h, 0, combined, offset, h.Length);
            Buffer.BlockCopy(salt, 0, combined, offset + h.Length, salt.Length);
        }

        h = CryptoHelpers.HashBytes(combined, HashIdentifier.MD5);
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

    public override int GenerateBlockKey(int blockNumber, byte[] secretKey, byte[] destination)
    {
        CryptoHelpers.HashBlockKey(secretKey, blockNumber, HashIdentifier.MD5, 16, destination, 16);
        return 16;
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

        using var cipher = CryptoHelpers.CreateCipher(CipherIdentifier.RC4, 0, 0, 0);
        using var transform = cipher.CreateDecryptor(blockKey, null);
        var decryptedVerifier = CryptoHelpers.DecryptBytes(transform, EncryptedVerifier, EncryptedVerifier.Length);
        var decryptedVerifierHash = CryptoHelpers.DecryptBytes(transform, EncryptedVerifierHash, EncryptedVerifierHash.Length);

        var verifierHash = CryptoHelpers.HashBytes(decryptedVerifier, HashIdentifier.MD5);
        for (var i = 0; i < 16; ++i)
        {
            if (decryptedVerifierHash[i] != verifierHash[i])
                return false;
        }

        return true;
    }
}
