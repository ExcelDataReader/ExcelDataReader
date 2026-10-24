#nullable enable

using ExcelDataReader.Core.CompoundFormat;
using ExcelDataReader.Core.OfficeCrypto;

namespace ExcelDataReader.Tests;

public class EncryptionTests
{
    [TestCase("")]
    [TestCase("password")]
    [TestCase("\u00e5\u4e2d\ud83d\ude00")]
    [TestCase("\ud800x\udc00")]
    public void SaltedPasswordEncodingPreservesUnicodeAndFallback(string password)
    {
        byte[] salt = [0, 255, 17];
        var expected = salt.Concat(System.Text.Encoding.Unicode.GetBytes(password)).ToArray();

        Assert.That(CryptoHelpers.EncodeSaltedPassword(password, salt), Is.EqualTo(expected));
        Assert.That(salt, Is.EqualTo(new byte[] { 0, 255, 17 }));
        Assert.That(CryptoHelpers.EncodeSaltedPassword(password, []), Is.EqualTo(System.Text.Encoding.Unicode.GetBytes(password)));
    }

    [TestCase("agile_AES128_SHA1_CBC_pwd_password.xlsx", "password")]
    [TestCase("agile_AES192_SHA512_CBC_pwd_password.xlsx", "password")]
    [TestCase("standard_AES256_SHA1_ECB_pwd_password.xlsx", "password")]
    [TestCase("Issue289.xlsx", "VelvetSweatshop")]
    public void VerifiedKeyMatchesIndependentDerivationWithoutRetainingFailedAttempts(string file, string password)
    {
        using var source = Configuration.GetTestWorkbook(file);
        var document = new CompoundDocument(source);
        var info = document.FindEntry("EncryptionInfo")!;
        var encryption = EncryptionInfo.Create(document.ReadStream(source, info.StreamFirstSector, (int)info.StreamSize, info.IsEntryMiniStream));

        Assert.That(encryption.TryVerifyPassword("wrongpassword", out var failedKey), Is.False);
        Assert.That(failedKey, Is.Null);
        Assert.That(encryption.TryVerifyPassword(password, out var key), Is.True);
        Assert.That(key, Is.EqualTo(encryption.GenerateSecretKey(password)));

        Assert.That(encryption.TryVerifyPassword("wrongpassword", out failedKey), Is.False);
        Assert.That(failedKey, Is.Null);
        Assert.That(encryption.TryVerifyPassword(password, out var repeatedKey), Is.True);
        Assert.That(repeatedKey, Is.EqualTo(key));

        var package = document.FindEntry("EncryptedPackage")!;
        using var decrypted = encryption.CreateEncryptedPackageStream(
            new CompoundStream(document, source, package.StreamFirstSector, (int)package.StreamSize, package.IsEntryMiniStream, true), key!);
        var signature = new byte[4];
        Assert.That(decrypted.Read(signature, 0, signature.Length), Is.EqualTo(signature.Length));
        Assert.That(signature, Is.EqualTo(new byte[] { 0x50, 0x4b, 0x03, 0x04 }));
        if (!decrypted.CanSeek)
            decrypted.CopyTo(Stream.Null);
    }

    [TestCase("Issue289.xlsx", null, "aaaaaaa")]
    [TestCase("Issue289.xlsx", "wrongpassword", "aaaaaaa")]
    [TestCase("protectedsheet-xxx.xls", null, "x")]
    [TestCase("protectedsheet-xxx.xls", "wrongpassword", "x")]
    public void MagicPasswordStillTakesPrecedenceOverSuppliedPassword(string file, string? password, string expected)
    {
        using var reader = ExcelReaderFactory.CreateReader(
            Configuration.GetTestWorkbook(file),
            new ExcelReaderConfiguration { Password = password });

        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo(expected));
            reader.Reset();
        }
    }
}