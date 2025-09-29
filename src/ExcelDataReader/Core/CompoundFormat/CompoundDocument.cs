﻿using System.Text;
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
        if (Header.ByteOrder != 0xFFFE && Header.ByteOrder != 0xFFFF) // Some broken xls files uses 0xFFFF
            throw new HeaderException(Errors.ErrorHeaderOrder);

        var difSectorChain = ReadDifSectorChain(reader);
        SectorTable = ReadSectorTable(reader, difSectorChain);

        var miniChain = GetSectorChain(Header.MiniFatFirstSector, SectorTable);
        MiniSectorTable = ReadSectorTable(reader, miniChain);

        var directoryChain = GetSectorChain(Header.RootDirectoryEntryStart, SectorTable);
        var bytes = ReadStream(stream, directoryChain, directoryChain.Count * Header.SectorSize);
        ReadDirectoryEntries(bytes);
    }

    internal CompoundHeader Header { get; }

    internal List<uint> SectorTable { get; }

    internal List<uint> MiniSectorTable { get; }

    internal CompoundDirectoryEntry RootEntry { get; set; }

    internal List<CompoundDirectoryEntry> Entries { get; set; }

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

    internal CompoundDirectoryEntry FindEntry(params string[] entryNames)
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

    internal byte[] ReadStream(Stream stream, List<uint> sectors, int length)
    {
        using var cfb = new CompoundStream(this, stream, sectors, length, true);
        var bytes = new byte[length];
        cfb.ReadAtLeast(bytes, 0, length);
        return bytes;
    }

    private static CompoundHeader ReadHeader(BinaryReader reader)
    {
        var result = new CompoundHeader();
        result.Signature = reader.ReadUInt64();
        result.ClassId = new Guid(reader.ReadBytes(16));
        result.Version = reader.ReadUInt16();
        result.DllVersion = reader.ReadUInt16();
        result.ByteOrder = reader.ReadUInt16();
        result.SectorSizeInPot = reader.ReadUInt16();
        result.MiniSectorSizeInPot = reader.ReadUInt16();
        reader.ReadBytes(6); // skip 6 unused bytes
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

    private void ReadDirectoryEntries(byte[] bytes)
    {
        try
        {
            Entries = [];
            using var stream = new MemoryStream(bytes);
            using var reader = new BinaryReader(stream);
            RootEntry = ReadDirectoryEntry(reader);
            Entries.Add(RootEntry);

            while (stream.Position < stream.Length)
            {
                var entry = ReadDirectoryEntry(reader);
                Entries.Add(entry);
            }
        }
        catch (EndOfStreamException ex)
        {
            throw new CompoundDocumentException(Errors.ErrorEndOfFile, ex);
        }
    }

    private CompoundDirectoryEntry ReadDirectoryEntry(BinaryReader reader)
    {
        var result = new CompoundDirectoryEntry();
        var name = reader.ReadBytes(64);
        var nameLength = reader.ReadUInt16();

        if (nameLength > 0)
        {
            nameLength = Math.Min((ushort)64, nameLength);
            result.EntryName = Encoding.Unicode.GetString(name, 0, nameLength).TrimEnd('\0');
        }

        result.EntryType = (STGTY)reader.ReadByte();
        result.EntryColor = (DECOLOR)reader.ReadByte();
        result.LeftSiblingSid = reader.ReadUInt32();
        result.RightSiblingSid = reader.ReadUInt32();
        result.ChildSid = reader.ReadUInt32();
        result.ClassId = new Guid(reader.ReadBytes(16));
        result.UserFlags = reader.ReadUInt32();
        result.CreationTime = ReadFileTime(reader);
        result.LastWriteTime = ReadFileTime(reader);
        result.StreamFirstSector = reader.ReadUInt32();
        result.StreamSize = reader.ReadUInt32();
        result.PropType = reader.ReadUInt32();
        result.IsEntryMiniStream = result.StreamSize < Header.MiniStreamCutoff;
        return result;

        static DateTime ReadFileTime(BinaryReader reader)
        {
            var d = reader.ReadInt64();
            if (d < 0 || d > SafeFileTimeMaxDate)
            {
                d = 0;
            }

            return DateTime.FromFileTime(d);
        }
    }

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
                    var difContent = ReadSectorAsUInt32s(reader, difSector);
                    difSectorChain.AddRange(difContent.GetRange(0, difContent.Count - 1));

                    // The DIFAT sectors are linked together by the "Next DIFAT Sector Location" in each DIFAT sector:
                    difSector = difContent[difContent.Count - 1];
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
                var result = ReadSectorAsUInt32s(reader, sector);
                sectorTable.AddRange(result);
            }
        }
        catch (EndOfStreamException ex)
        {
            throw new CompoundDocumentException(Errors.ErrorEndOfFile, ex);
        }

        TrimSectorChain(sectorTable, FATMARKERS.FAT_FreeSpace);

        return sectorTable;
    }

    private List<uint> ReadSectorAsUInt32s(BinaryReader reader, uint sector)
    {
        var result = new List<uint>(Header.SectorSize / 4);
        reader.BaseStream.Seek(GetSectorOffset(sector), SeekOrigin.Begin);
        for (var i = 0; i < Header.SectorSize / 4; ++i)
        {
            var value = reader.ReadUInt32();
            result.Add(value);
        }

        return result;
    }
}
