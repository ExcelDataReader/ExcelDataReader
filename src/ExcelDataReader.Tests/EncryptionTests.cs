#nullable enable

using ExcelDataReader.Core;
using ExcelDataReader.Core.CompoundFormat;
using ExcelDataReader.Core.OfficeCrypto;

namespace ExcelDataReader.Tests;

public class EncryptionTests
{
    [TestCase(false)]
    [TestCase(true)]
    public void DecryptionOnlyAlgorithmsRejectEncryptionAndKeyGeneration(bool xor)
    {
        using System.Security.Cryptography.SymmetricAlgorithm algorithm = xor ? new XorManaged() : new RC4Managed();

        Assert.Throws<NotSupportedException>(() => algorithm.CreateEncryptor(new byte[16], null));
        Assert.Throws<NotSupportedException>(algorithm.GenerateIV);
        Assert.Throws<NotSupportedException>(algorithm.GenerateKey);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void ReadOnlyCompoundAndAgileStreamsRejectWriting(bool decrypt)
    {
        using var source = Configuration.GetTestWorkbook("agile_AES128_SHA1_CBC_pwd_password.xlsx");
        var document = new CompoundDocument(source);
        var package = document.FindEntry("EncryptedPackage")!;
        using var input = new CompoundStream(document, source, package.StreamFirstSector, (int)package.StreamSize, package.IsEntryMiniStream, true);
        Stream stream = input;
        if (decrypt)
        {
            var info = document.FindEntry("EncryptionInfo")!;
            var encryption = EncryptionInfo.Create(document.ReadStream(source, info.StreamFirstSector, (int)info.StreamSize, info.IsEntryMiniStream));
            stream = encryption.CreateEncryptedPackageStream(input, encryption.GenerateSecretKey("password"));
        }

        using (stream)
        {
            Assert.That(stream.CanWrite, Is.False);
            Assert.Throws<NotSupportedException>(() => stream.Write(new byte[1], 0, 1));
            Assert.Throws<NotSupportedException>(() => stream.SetLength(0));
            Assert.Throws<NotSupportedException>(() => stream.WriteByte(0));
        }
    }

    [TestCase(false)]
    [TestCase(true)]
    public void CleanupAttemptsEveryResourceAndPreservesFirstFailure(bool operationFailed)
    {
        var first = new DisposalTrackingStream([]) { DisposeError = new IOException("First cleanup failure.") };
        var second = new DisposalTrackingStream([]) { DisposeError = new IOException("Second cleanup failure.") };
        var original = new InvalidOperationException("Operation failed.");

        void FailAndCleanup()
        {
            try
            {
                throw original;
            }
            catch (Exception error)
            {
                ResourceCleanup.DisposeAll(error, first, second);
                throw;
            }
        }

        if (operationFailed)
            Assert.That(Assert.Throws<InvalidOperationException>(FailAndCleanup), Is.SameAs(original));
        else
            Assert.That(Assert.Throws<IOException>(() => ResourceCleanup.DisposeAll(null, first, second)), Is.SameAs(first.DisposeError));

        Assert.That(first.DisposeCount, Is.EqualTo(1));
        Assert.That(second.DisposeCount, Is.EqualTo(1));
    }

    [TestCase(3)]
    [TestCase(4)]
    public void ExtensibleEncryptionIsRecognizedButNotSupported(byte majorVersion)
    {
        Assert.Throws<NotSupportedException>(() => EncryptionInfo.Create([majorVersion, 0, 3, 0]));
    }

    [Test]
    public void StandardAesEncryptionRejectsBlockKeyGeneration()
    {
        using var source = Configuration.GetTestWorkbook("standard_AES256_SHA1_ECB_pwd_password.xlsx");
        var document = new CompoundDocument(source);
        var info = document.FindEntry("EncryptionInfo")!;
        var encryption = EncryptionInfo.Create(document.ReadStream(source, info.StreamFirstSector, (int)info.StreamSize, info.IsEntryMiniStream));

        Assert.Throws<InvalidOperationException>(() => encryption.GenerateBlockKey(0, new byte[32], new byte[CryptoHelpers.MaxHashSize]));
    }

    [TestCase(0)]
    [TestCase(4)]
    [TestCase(-1)]
    public void StandardPackageDisposalHandlesUnreadPartialAndCompleteReads(int readCount)
    {
        using var source = Configuration.GetTestWorkbook("standard_AES256_SHA1_ECB_pwd_password.xlsx");
        var document = new CompoundDocument(source);
        var info = document.FindEntry("EncryptionInfo")!;
        var encryption = EncryptionInfo.Create(document.ReadStream(source, info.StreamFirstSector, (int)info.StreamSize, info.IsEntryMiniStream));
        var package = document.FindEntry("EncryptedPackage")!;
        using var input = new DisposalTrackingStream(document.ReadStream(source, package.StreamFirstSector, (int)package.StreamSize, package.IsEntryMiniStream));
        var decrypted = encryption.CreateEncryptedPackageStream(input, encryption.GenerateSecretKey("password"));
        try
        {
            if (readCount < 0)
                decrypted.CopyTo(Stream.Null);
            else if (readCount > 0)
                Assert.That(decrypted.Read(new byte[readCount], 0, readCount), Is.EqualTo(readCount));
        }
        finally
        {
            Assert.DoesNotThrow(decrypted.Dispose);
        }

        Assert.That(input.CanRead, Is.False);
        Assert.Throws<ObjectDisposedException>(() => _ = decrypted.Read(new byte[1], 0, 1));
        Assert.DoesNotThrow(decrypted.Dispose);
        Assert.That(input.DisposeCount, Is.EqualTo(1));
    }

    [Test]
    public void StandardPackageDisposalPropagatesFailureWithoutRetryingDisposedResources()
    {
        using var source = Configuration.GetTestWorkbook("standard_AES256_SHA1_ECB_pwd_password.xlsx");
        var document = new CompoundDocument(source);
        var info = document.FindEntry("EncryptionInfo")!;
        var encryption = EncryptionInfo.Create(document.ReadStream(source, info.StreamFirstSector, (int)info.StreamSize, info.IsEntryMiniStream));
        var package = document.FindEntry("EncryptedPackage")!;
        var input = new DisposalTrackingStream(document.ReadStream(source, package.StreamFirstSector, (int)package.StreamSize, package.IsEntryMiniStream))
        {
            DisposeError = new IOException("Disposal failed.")
        };
        var decrypted = encryption.CreateEncryptedPackageStream(input, encryption.GenerateSecretKey("password"));
        Assert.That(decrypted.Read(new byte[4], 0, 4), Is.EqualTo(4));

        Assert.That(Assert.Throws<IOException>(decrypted.Dispose), Is.SameAs(input.DisposeError));
        Assert.That(input.CanRead, Is.False);
        Assert.Throws<ObjectDisposedException>(() => _ = decrypted.Read(new byte[1], 0, 1));
        Assert.DoesNotThrow(decrypted.Dispose);
        Assert.That(input.DisposeCount, Is.EqualTo(1));
    }

    [TestCase("agile_AES128_SHA1_CBC_pwd_password.xlsx")]
    [TestCase("agile_AES192_SHA512_CBC_pwd_password.xlsx")]
    [TestCase("agile_DESede_SHA384_CBC_pwd_password.xlsx")]
    [TestCase("agile_DES_MD5_CBC_pwd_password.xlsx")]
    [TestCase("agile_RC2_SHA1_CBC_pwd_password.xlsx")]
    public void AgileSegmentsPreservePlaintextAcrossReadsAndSeeks(string file)
    {
        using var source = Configuration.GetTestWorkbook(file);
        var document = new CompoundDocument(source);
        var info = document.FindEntry("EncryptionInfo")!;
        var encryption = (AgileEncryption)EncryptionInfo.Create(document.ReadStream(source, info.StreamFirstSector, (int)info.StreamSize, info.IsEntryMiniStream));
        var key = encryption.GenerateSecretKey("password");
        var expected = new byte[(3 * 4096) - 13];
        new Random(42).NextBytes(expected);
        using var package = new MemoryStream();
        package.Write(BitConverter.GetBytes((long)expected.Length), 0, 8);
        using (var cipher = encryption.CreateCipher())
        {
            cipher.Padding = System.Security.Cryptography.PaddingMode.Zeros;
            for (int segment = 0; segment < 3; segment++)
            {
                using var encryptor = cipher.CreateEncryptor(key, encryption.GenerateBlockKey(segment, encryption.SaltValue));
                var bytes = encryptor.TransformFinalBlock(expected, segment * 4096, Math.Min(4096, expected.Length - (segment * 4096)));
                package.Write(bytes, 0, bytes.Length);
            }
        }

        package.Position = 0;
        using var decrypted = encryption.CreateEncryptedPackageStream(package, key);
        var actual = new byte[expected.Length];
        int offset = 0;
        while (offset < actual.Length)
        {
            int count = Math.Min(1337, actual.Length - offset);
            Assert.That(decrypted.Read(actual, offset, count), Is.EqualTo(count));
            offset += count;
        }

        Assert.That(actual, Is.EqualTo(expected));
        foreach (int position in new[] { 4095, 8192, 1, 4096, 0 })
        {
            Assert.That(decrypted.Seek(position, SeekOrigin.Begin), Is.EqualTo(position));
            var bytes = new byte[17];
            Assert.That(decrypted.Read(bytes, 0, bytes.Length), Is.EqualTo(bytes.Length));
            Assert.That(bytes, Is.EqualTo(expected.Skip(position).Take(bytes.Length).ToArray()));
        }
    }

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

    [Test]
    public void AgilePackageDisposalPropagatesFailureWithoutRetryingDisposedResources()
    {
        using var source = Configuration.GetTestWorkbook("agile_AES128_SHA1_CBC_pwd_password.xlsx");
        var document = new CompoundDocument(source);
        var info = document.FindEntry("EncryptionInfo")!;
        var encryption = EncryptionInfo.Create(document.ReadStream(source, info.StreamFirstSector, (int)info.StreamSize, info.IsEntryMiniStream));
        var package = document.FindEntry("EncryptedPackage")!;
        var input = new DisposalTrackingStream(document.ReadStream(source, package.StreamFirstSector, (int)package.StreamSize, package.IsEntryMiniStream))
        {
            DisposeError = new IOException("Disposal failed.")
        };
        var decrypted = encryption.CreateEncryptedPackageStream(input, encryption.GenerateSecretKey("password"));

        Assert.That(Assert.Throws<IOException>(decrypted.Dispose), Is.SameAs(input.DisposeError));
        Assert.That(input.CanRead, Is.False);
        Assert.DoesNotThrow(decrypted.Dispose);
        Assert.That(input.DisposeCount, Is.EqualTo(1));
    }

    private sealed class DisposalTrackingStream(byte[] bytes) : MemoryStream(bytes)
    {
        public IOException? DisposeError { get; set; }

        public int DisposeCount { get; private set; }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing)
            {
                DisposeCount++;
                if (DisposeError != null)
                    throw DisposeError;
            }
        }
    }
}