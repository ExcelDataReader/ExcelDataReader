using System.Text;
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
using System.Buffers.Binary;
#endif
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.CompoundFormat;

internal sealed class CompoundDocument
{
    public CompoundDocument(Stream stream)
    {
        var reader = new BinaryReader(stream);

        Header = ReadHeader(reader);

        if (!Header.IsSignatureValid)
            throw new HeaderException(Errors.ErrorHeaderSignature);
        if (Header.ByteOrder != 0xFFFE && Header.ByteOrder != 0xFFFF && Header.ByteOrder != 0xFEFF) // Some broken xls files use 0xFFFF or 0xFEFF
            throw new HeaderException(Errors.ErrorHeaderOrder);

        var difSectorChain = ReadDifSectorChain(reader);
        SectorTable = ReadSectorTable(reader, difSectorChain);

        var miniChain = GetSectorChain(Header.MiniFatFirstSector, SectorTable);
        MiniSectorTable = ReadSectorTable(reader, miniChain);

        var directoryChain = GetSectorChain(Header.RootDirectoryEntryStart, SectorTable);
        using var directoryStream = new CompoundStream(this, stream, directoryChain, directoryChain.Count * Header.SectorSize, true);
        ReadDirectoryEntries(directoryStream);
    }

    internal CompoundHeader Header { get; }

    internal List<uint> SectorTable { get; }

    internal List<uint> MiniSectorTable { get; }

    internal CompoundDirectoryEntry RootEntry { get; set; } = null!;

    internal List<CompoundDirectoryEntry> Entries { get; set; } = [];

    // NOTE: DateTime.MaxValue.ToFileTime() fails on Unity in timezones with DST and +~6h offset, like Sidney Australia
    private static long SafeFileTimeMaxDate { get; } = DateTime.MaxValue.ToFileTimeUtc();

    internal static List<uint> GetSectorChain(uint sector, List<uint> sectorTable)
    {
        List<uint> chain = [];
        while (sector != (uint)FATMARKERS.FAT_EndOfChain)
        {
            chain.Add(sector);
            sector = GetNextSector(sector, sectorTable);

            if (chain.Contains(sector))
            {
                throw new CompoundDocumentException(Errors.ErrorCyclicSectorChain);
            }
        }

        TrimSectorChain(chain, FATMARKERS.FAT_FreeSpace);
        return chain;
    }

    internal static bool IsCompoundDocument(byte[] probe)
    {
        return BitConverter.ToUInt64(probe, 0) == 0xE11AB1A1E011CFD0;
    }

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    // params ReadOnlySpan<T> avoids the implicit array allocation on modern runtimes.
    internal CompoundDirectoryEntry? FindEntry(params ReadOnlySpan<string> entryNames)
#else
    internal CompoundDirectoryEntry? FindEntry(params string[] entryNames)
#endif
    {
        foreach (var e in Entries)
        {
            foreach (var entryName in entryNames)
            {
                if (string.Equals(e.EntryName, entryName, StringComparison.OrdinalIgnoreCase))
                    return e;
            }
        }

        return null;
    }

    internal long GetMiniSectorOffset(uint sector)
    {
        return Header.MiniSectorSize * sector;
    }

    internal long GetSectorOffset(uint sector)
    {
        return 512 + Header.SectorSize * sector;
    }

    /// <summary>
    /// Reads bytes from a regular or mini stream.
    /// </summary>
    internal byte[] ReadStream(Stream stream, uint baseSector, int length, bool isMini)
    {
        using var cfb = new CompoundStream(this, stream, baseSector, length, isMini, true);
        var bytes = new byte[length];
        cfb.ReadAtLeast(bytes, 0, length);
        return bytes;
    }

    private static CompoundHeader ReadHeader(BinaryReader reader)
    {
        var result = new CompoundHeader();
        result.Signature = reader.ReadUInt64();
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        Span<byte> guid = stackalloc byte[16];
        if (reader.BaseStream.ReadAtLeast(guid) != guid.Length)
            throw new EndOfStreamException();
        result.ClassId = new Guid(guid);
#else
        result.ClassId = new Guid(reader.ReadBytes(16));
#endif
        result.Version = reader.ReadUInt16();
        result.DllVersion = reader.ReadUInt16();
        result.ByteOrder = reader.ReadUInt16();
        result.SectorSizeInPot = reader.ReadUInt16();
        result.MiniSectorSizeInPot = reader.ReadUInt16();
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        reader.ReadUInt32(); // skip 6 unused bytes
        reader.ReadUInt16();
#else
        reader.ReadBytes(6); // skip 6 unused bytes
#endif
        result.DirectorySectorCount = reader.ReadInt32();
        result.FatSectorCount = reader.ReadInt32();
        result.RootDirectoryEntryStart = reader.ReadUInt32();
        result.TransactionSignature = reader.ReadUInt32();
        result.MiniStreamCutoff = reader.ReadUInt32();
        result.MiniFatFirstSector = reader.ReadUInt32();
        result.MiniFatSectorCount = reader.ReadInt32();
        result.DifFirstSector = reader.ReadUInt32();
        result.DifSectorCount = reader.ReadInt32();

        List<uint> chain = [];
        for (int i = 0; i < 109; ++i)
        {
            chain.Add(reader.ReadUInt32());
        }

        result.First109DifSectorChain = chain;

        return result;
    }

    private static void TrimSectorChain(List<uint> chain, FATMARKERS marker)
    {
        while (chain.Count > 0 && chain[chain.Count - 1] == (uint)marker)
        {
            chain.RemoveAt(chain.Count - 1);
        }
    }

    private static uint GetNextSector(uint sector, List<uint> sectorTable)
    {
        if (sector < sectorTable.Count)
        {
            return sectorTable[(int)sector];
        }

        return (uint)FATMARKERS.FAT_EndOfChain;
    }

    private static DateTime ReadFileTime(long d)
    {
        if (d < 0 || d > SafeFileTimeMaxDate)
        {
            d = 0;
        }

        return DateTime.FromFileTime(d);
    }

    private void ReadDirectoryEntries(Stream stream)
    {
        try
        {
            Entries = [];
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
            Span<byte> entryBuffer = stackalloc byte[128];
#else
            var entryBuffer = new byte[128];
#endif
            RootEntry = ReadDirectoryEntry(stream, entryBuffer);
            Entries.Add(RootEntry);

            while (stream.Position < stream.Length)
            {
                var entry = ReadDirectoryEntry(stream, entryBuffer);
                Entries.Add(entry);
            }
        }
        catch (EndOfStreamException ex)
        {
            throw new CompoundDocumentException(Errors.ErrorEndOfFile, ex);
        }
    }

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
    private CompoundDirectoryEntry ReadDirectoryEntry(Stream stream, Span<byte> buffer)
    {
        if (stream.ReadAtLeast(buffer) != buffer.Length)
            throw new EndOfStreamException();

        var result = new CompoundDirectoryEntry();
        var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(buffer[64..]);

        if (nameLength > 0)
        {
            nameLength = Math.Min((ushort)64, nameLength);
            result.EntryName = Encoding.Unicode.GetString(buffer[..nameLength]).TrimEnd('\0');
        }

        result.EntryType = (STGTY)buffer[66];
        result.EntryColor = (DECOLOR)buffer[67];
        result.LeftSiblingSid = BinaryPrimitives.ReadUInt32LittleEndian(buffer[68..]);
        result.RightSiblingSid = BinaryPrimitives.ReadUInt32LittleEndian(buffer[72..]);
        result.ChildSid = BinaryPrimitives.ReadUInt32LittleEndian(buffer[76..]);
        result.ClassId = new Guid(buffer.Slice(80, 16));
        result.UserFlags = BinaryPrimitives.ReadUInt32LittleEndian(buffer[96..]);
        result.CreationTime = ReadFileTime(BinaryPrimitives.ReadInt64LittleEndian(buffer[100..]));
        result.LastWriteTime = ReadFileTime(BinaryPrimitives.ReadInt64LittleEndian(buffer[108..]));
        result.StreamFirstSector = BinaryPrimitives.ReadUInt32LittleEndian(buffer[116..]);
        result.StreamSize = BinaryPrimitives.ReadUInt32LittleEndian(buffer[120..]);
        result.PropType = BinaryPrimitives.ReadUInt32LittleEndian(buffer[124..]);
        result.IsEntryMiniStream = result.StreamSize < Header.MiniStreamCutoff;
        return result;
    }
#else
    private CompoundDirectoryEntry ReadDirectoryEntry(Stream stream, byte[] buffer)
    {
        if (stream.ReadAtLeast(buffer, 0, buffer.Length) != buffer.Length)
            throw new EndOfStreamException();

        var result = new CompoundDirectoryEntry();
        var nameLength = BitConverter.ToUInt16(buffer, 64);

        if (nameLength > 0)
        {
            nameLength = Math.Min((ushort)64, nameLength);
            result.EntryName = Encoding.Unicode.GetString(buffer, 0, nameLength).TrimEnd('\0');
        }

        result.EntryType = (STGTY)buffer[66];
        result.EntryColor = (DECOLOR)buffer[67];
        result.LeftSiblingSid = BitConverter.ToUInt32(buffer, 68);
        result.RightSiblingSid = BitConverter.ToUInt32(buffer, 72);
        result.ChildSid = BitConverter.ToUInt32(buffer, 76);
        result.ClassId = new Guid(
            BitConverter.ToInt32(buffer, 80),
            BitConverter.ToInt16(buffer, 84),
            BitConverter.ToInt16(buffer, 86),
            buffer[88],
            buffer[89],
            buffer[90],
            buffer[91],
            buffer[92],
            buffer[93],
            buffer[94],
            buffer[95]);
        result.UserFlags = BitConverter.ToUInt32(buffer, 96);
        result.CreationTime = ReadFileTime(BitConverter.ToInt64(buffer, 100));
        result.LastWriteTime = ReadFileTime(BitConverter.ToInt64(buffer, 108));
        result.StreamFirstSector = BitConverter.ToUInt32(buffer, 116);
        result.StreamSize = BitConverter.ToUInt32(buffer, 120);
        result.PropType = BitConverter.ToUInt32(buffer, 124);
        result.IsEntryMiniStream = result.StreamSize < Header.MiniStreamCutoff;
        return result;
    }
#endif

    /// <summary>
    /// The header contains the first 109 DIF entries. If there are any more, read from a separate stream.
    /// </summary>
    private List<uint> ReadDifSectorChain(BinaryReader reader)
    {
        // Read the DIF chain sectors directly, can't use ReadStream yet because it depends on the DIF chain
        var difSectorChain = new List<uint>(Header.First109DifSectorChain);
        if (Header.DifFirstSector != (uint)FATMARKERS.FAT_EndOfChain)
        {
            try
            {
                var difSector = Header.DifFirstSector;
                for (var i = 0; i < Header.DifSectorCount; ++i)
                {
                    ReadSectorAsUInt32s(reader, difSector, difSectorChain, Header.SectorSize / 4 - 1);

                    // The DIFAT sectors are linked together by the "Next DIFAT Sector Location" in each DIFAT sector:
                    difSector = reader.ReadUInt32();
                }
            }
            catch (EndOfStreamException ex)
            {
                throw new CompoundDocumentException(Errors.ErrorEndOfFile, ex);
            }
        }

        TrimSectorChain(difSectorChain, FATMARKERS.FAT_FreeSpace);

        // A special value of ENDOFCHAIN (0xFFFFFFFE) is stored in the "Next DIFAT Sector Location" field of the
        // last DIFAT sector, or in the header when no DIFAT sectors are needed.
        TrimSectorChain(difSectorChain, FATMARKERS.FAT_EndOfChain);

        return difSectorChain;
    }

    private List<uint> ReadSectorTable(BinaryReader reader, List<uint> chain)
    {
        var sectorTable = new List<uint>(Header.SectorSize / 4 * chain.Count);
        try
        {
            foreach (var sector in chain)
            {
                ReadSectorAsUInt32s(reader, sector, sectorTable, Header.SectorSize / 4);
            }
        }
        catch (EndOfStreamException ex)
        {
            throw new CompoundDocumentException(Errors.ErrorEndOfFile, ex);
        }

        TrimSectorChain(sectorTable, FATMARKERS.FAT_FreeSpace);

        return sectorTable;
    }

    private void ReadSectorAsUInt32s(BinaryReader reader, uint sector, List<uint> destination, int count)
    {
        reader.BaseStream.Seek(GetSectorOffset(sector), SeekOrigin.Begin);
        for (var i = 0; i < count; ++i)
        {
            var value = reader.ReadUInt32();
            destination.Add(value);
        }
    }
}