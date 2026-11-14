using ExcelDataReader.Core.CompoundFormat;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Tests;

public class CompoundDocumentTests
{
    [TestCase(1, 0)]
    [TestCase(2, 0)]
    [TestCase(31, 0)]
    [TestCase(32, 7)]
    [TestCase(33, 32)]
    [TestCase(64, 7)]
    public void SectorChainRejectsCycles(int length, int cycleStart)
    {
        var sectorTable = Enumerable.Range(1, length).Select(i => (uint)i).ToList();
        sectorTable[length - 1] = (uint)cycleStart;
        Assert.Throws<CompoundDocumentException>(() =>
            CompoundDocument.GetSectorChain(0, sectorTable));
    }

    [Test]
    public void SectorChainPreservesOrderAndEndMarkerBehavior()
    {
        const uint end = (uint)FATMARKERS.FAT_EndOfChain;
        const uint free = (uint)FATMARKERS.FAT_FreeSpace;
        List<uint> sectorTable = [3, end, 0, end, free];

        Assert.That(CompoundDocument.GetSectorChain(2, sectorTable), Is.EqualTo(new uint[] { 2, 0, 3 }));
        Assert.That(CompoundDocument.GetSectorChain(end, sectorTable), Is.Empty);
        Assert.That(CompoundDocument.GetSectorChain(free, sectorTable), Is.Empty);
        Assert.That(CompoundDocument.GetSectorChain(4, sectorTable), Is.EqualTo(new uint[] { 4 }));
        Assert.That(CompoundDocument.GetSectorChain(42, sectorTable), Is.EqualTo(new uint[] { 42 }));
    }
}
