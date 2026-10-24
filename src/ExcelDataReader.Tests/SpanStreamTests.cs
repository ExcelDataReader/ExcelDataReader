#if NET8_0_OR_GREATER
using ExcelDataReader.Core.CompoundFormat;
using ExcelDataReader.Core.OfficeCrypto;
using ExcelDataReader.Misc;
using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Tests;

public class SpanStreamTests
{
    [Test]
    public void CompoundSpanDecodingPreservesGuidsAndOwnedNames()
    {
        using var source = Configuration.GetTestWorkbook("Issue532_Mulcells.xls");
        var original = new CompoundDocument(source);
        source.Position = 0;
        using var modified = new MemoryStream();
        source.CopyTo(modified);
        var expected = new Guid("01234567-89ab-cdef-0123-456789abcdef");
        modified.Position = original.GetSectorOffset(original.Header.RootDirectoryEntryStart) + 80;
        var bytes = expected.ToByteArray();
        modified.Write(bytes, 0, bytes.Length);
        modified.Position = 8;
        modified.Write(bytes, 0, bytes.Length);
        modified.Position = 34;
        modified.Write(new byte[] { 1, 2, 3, 4, 5, 6 }, 0, 6);
        modified.Position = 0;

        var document = new CompoundDocument(modified);

        Assert.That(document.RootEntry.ClassId, Is.EqualTo(expected));
        Assert.That(document.Header.ClassId, Is.EqualTo(expected));
        Assert.That(document.Header.FatSectorCount, Is.EqualTo(original.Header.FatSectorCount));
        Assert.That(document.Entries.Select(entry => entry.EntryName), Is.EqualTo(original.Entries.Select(entry => entry.EntryName)));
    }

    [TestCase("BigFormatted.xls", "Workbook")]
    [TestCase("agile_AES128_SHA1_CBC_pwd_password.xlsx", "EncryptionInfo")]
    public void CompoundSpanReadsMatchArrayReadsAcrossSectorsAndSeeks(string file, string entryName)
    {
        using var source = Configuration.GetTestWorkbook(file);
        var document = new CompoundDocument(source);
        var entry = document.FindEntry(entryName)!;
        using var arrayStream = new CompoundStream(document, source, entry.StreamFirstSector, (int)entry.StreamSize, entry.IsEntryMiniStream, true);
        using var spanStream = new CompoundStream(document, source, entry.StreamFirstSector, (int)entry.StreamSize, entry.IsEntryMiniStream, true);

        AssertMatchingReads(arrayStream, spanStream);
        arrayStream.Seek(3, SeekOrigin.Begin);
        spanStream.Seek(3, SeekOrigin.Begin);
        AssertMatchingReads(arrayStream, spanStream, 3);
        Assert.That(spanStream.Read(Span<byte>.Empty), Is.Zero);
        Assert.That(spanStream.Read(new byte[1].AsSpan()), Is.Zero);
    }

    [TestCase("agile_AES128_SHA1_CBC_pwd_password.xlsx")]
    [TestCase("standard_AES128_SHA1_ECB_pwd_password.xlsx")]
    public void EncryptedSpanReadsMatchArrayReadsAcrossSegments(string file)
    {
        using var source = Configuration.GetTestWorkbook(file);
        var document = new CompoundDocument(source);
        var info = document.FindEntry("EncryptionInfo")!;
        var encryption = EncryptionInfo.Create(document.ReadStream(source, info.StreamFirstSector, (int)info.StreamSize, info.IsEntryMiniStream));
        var key = encryption.GenerateSecretKey("password");
        var package = document.FindEntry("EncryptedPackage")!;
        using var arrayStream = encryption.CreateEncryptedPackageStream(
            new CompoundStream(document, source, package.StreamFirstSector, (int)package.StreamSize, package.IsEntryMiniStream, true), key);
        using var spanStream = encryption.CreateEncryptedPackageStream(
            new CompoundStream(document, source, package.StreamFirstSector, (int)package.StreamSize, package.IsEntryMiniStream, true), key);

        AssertMatchingReads(arrayStream, spanStream);
        if (arrayStream.CanSeek)
        {
            Assert.That(arrayStream.Read(new byte[1], 0, 1), Is.Zero);
            Assert.That(spanStream.Read(new byte[1].AsSpan()), Is.Zero);
            Assert.That(spanStream.Read(Span<byte>.Empty), Is.Zero);
            arrayStream.Seek(-1, SeekOrigin.End);
            spanStream.Seek(-1, SeekOrigin.End);
            var expected = new byte[16];
            var actual = new byte[16];
            Assert.That(arrayStream.Read(expected, 0, expected.Length), Is.EqualTo(1));
            Assert.That(spanStream.Read(actual.AsSpan()), Is.EqualTo(1));
            Assert.That(actual, Is.EqualTo(expected));
            Assert.That(spanStream.Position, Is.EqualTo(spanStream.Length));
            arrayStream.Seek(3, SeekOrigin.Begin);
            spanStream.Seek(3, SeekOrigin.Begin);
            AssertMatchingReads(arrayStream, spanStream, 3);
        }
        else
        {
            // Standard encryption exposes decrypted block padding beyond its advertised length.
            var expected = new byte[1];
            var actual = new byte[1];
            int count;
            do
            {
                count = arrayStream.Read(expected, 0, 1);
                Assert.That(spanStream.Read(actual.AsSpan()), Is.EqualTo(count));
                Assert.That(actual, Is.EqualTo(expected));
            }
            while (count != 0);
        }
    }

    [Test]
    public void LeaveOpenForwardsSpanReadsWithoutArrayBridge()
    {
        using var source = new SpanOnlyStream();
        using (var wrapper = new LeaveOpenStream(source))
        {
            Span<byte> bytes = stackalloc byte[3];
            Assert.That(wrapper.Read(bytes), Is.EqualTo(3));
            Assert.That(bytes.ToArray(), Is.EqualTo(new byte[] { 1, 2, 3 }));
        }

        Assert.That(source.CanRead, Is.True);
    }

    private static void AssertMatchingReads(Stream arrayStream, Stream spanStream, long start = 0)
    {
        var expected = new byte[4101];
        var actual = new byte[4101];
        int[] sizes = [1, 63, 511, 4097];
        int step = 0;
        long position = start;
        while (position < arrayStream.Length)
        {
            int size = (int)Math.Min(sizes[step++ % sizes.Length], arrayStream.Length - position);
            int count = arrayStream.Read(expected, 2, size);
            int spanCount = spanStream.Read(actual.AsSpan(2, size));
            Assert.That(spanCount, Is.EqualTo(count));
            Assert.That(count, Is.GreaterThan(0));
            Assert.That(actual, Is.EqualTo(expected));
            position += count;
        }
    }

    private sealed class SpanOnlyStream : MemoryStream
    {
        public SpanOnlyStream()
        {
            Write(new byte[] { 1, 2, 3 }, 0, 3);
            Position = 0;
        }

        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Unexpected array read.");

        public override int Read(Span<byte> buffer)
        {
            int count = (int)Math.Min(buffer.Length, Length - Position);
            GetBuffer().AsSpan((int)Position, count).CopyTo(buffer);
            Position += count;
            return count;
        }
    }
}
#endif