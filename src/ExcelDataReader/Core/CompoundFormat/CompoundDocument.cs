using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.CompoundFormat
{
    internal class CompoundDocument
    {
        public CompoundDocument(Stream stream)
        {
            var reader = new BinaryReader(stream);

            Header = ReadHeader(reader);

            if (!Header.IsSignatureValid)
                throw new HeaderException(Errors.ErrorHeaderSignature);
            if (Header.ByteOrder != 0xFFFE && Header.ByteOrder != 0xFFFF) // Some broken xls files uses 0xFFFF
                throw new FormatException(Errors.ErrorHeaderOrder);

            var difSectorChain = ReadDifSectorChain(reader);
            SectorTable = ReadSectorTable(reader, difSectorChain);

            var miniChain = GetSectorChain(Header.MiniFatFirstSector, SectorTable);
            MiniSectorTable = ReadSectorTable(reader, miniChain);

            var directoryChain = GetSectorChain(Header.RootDirectoryEntryStart, SectorTable);
            var bytes = ReadRegularStream(stream, directoryChain, directoryChain.Count * Header.SectorSize);
            ReadDirectoryEntries(bytes);
        }

        internal CompoundHeader Header { get; }

        internal List<uint> SectorTable { get; }

        internal List<uint> MiniSectorTable { get; }

        internal CompoundDirectoryEntry RootEntry { get; set; }

        internal List<CompoundDirectoryEntry> Entries { get; set; }

        internal static bool IsCompoundDocument(byte[] probe)
        {
            return BitConverter.ToUInt64(probe, 0) == 0xE11AB1A1E011CFD0;
        }

        internal CompoundDirectoryEntry FindEntry(string entryName)
        {
            foreach (var e in Entries)
            {
                if (string.Equals(e.EntryName, entryName, StringComparison.CurrentCultureIgnoreCase))
                    return e;
            }

            return null;
        }

        /// <summary>
        /// Creates a Stream instance to read from the compound document.
        /// </summary>
        internal Stream CreateStream(Stream stream, uint baseSector, int length, bool isMini)
        {
            return new CompoundStream(this, stream, baseSector, length, isMini);
        }

        internal long GetMiniSectorOffset(uint sector)
        {
            return Header.MiniSectorSize * sector;
        }

        internal long GetSectorOffset(uint sector)
        {
            return 512 + Header.SectorSize * sector;
        }

        internal List<uint> GetSectorChain(uint sector, List<uint> sectorTable)
        {
            List<uint> chain = new List<uint>();
            while (sector != (uint)FATMARKERS.FAT_EndOfChain)
            {
                chain.Add(sector);
                sector = GetNextSector(sector, sectorTable);
            }

            TrimSectorChain(chain, FATMARKERS.FAT_FreeSpace);
            return chain;
        }

        /// <summary>
        /// Reads bytes from a regular or mini stream.
        /// </summary>
        internal byte[] ReadStream(Stream stream, uint baseSector, int length, bool isMini)
        {
            if (isMini)
            {
                return ReadMiniStream(stream, baseSector, length);
            }
            else
            {
                return ReadRegularStream(stream, baseSector, length);
            }
        }

        /// <summary>
        /// Reads bytes from the mini stream stored inside the Root Entry's stream.
        /// </summary>
        private byte[] ReadMiniStream(Stream stream, uint baseSector, int length)
        {
            var chain = GetSectorChain(baseSector, MiniSectorTable);
            var rootStreamChain = GetSectorChain(RootEntry.StreamFirstSector, SectorTable);

            var result = new byte[length];
            int resultOffset = 0;
            foreach (var sector in chain)
            {
                // Convert to sector+offset in the root stream
                var miniStreamOffset = (int)GetMiniSectorOffset(sector);

                var rootSector = rootStreamChain[miniStreamOffset / Header.SectorSize];
                var rootOffset = miniStreamOffset % Header.SectorSize;

                stream.Seek(GetSectorOffset(rootSector) + rootOffset, SeekOrigin.Begin);

                var chunkSize = Math.Min(length - resultOffset, Header.MiniSectorSize);
                if (stream.Read(result, resultOffset, chunkSize) < chunkSize)
                {
                    throw new InvalidOperationException("The excel file may be corrupt or truncated. We've read past the end of the file.");
                }

                resultOffset += chunkSize;
            }

            return result;
        }

        private byte[] ReadRegularStream(Stream stream, uint baseSector, int length)
        {
            var chain = GetSectorChain(baseSector, SectorTable);
            return ReadRegularStream(stream, chain, length);
        }

        private byte[] ReadRegularStream(Stream stream, List<uint> chain, int length)
        {
            var result = new byte[length];
            int resultOffset = 0;
            foreach (var sector in chain)
            {
                stream.Seek(GetSectorOffset(sector), SeekOrigin.Begin);

                var chunkSize = Math.Min(length - resultOffset, Header.SectorSize);
                if (stream.Read(result, resultOffset, chunkSize) < chunkSize)
                {
                    throw new InvalidOperationException("The excel file may be corrupt or truncated. We've read past the end of the file.");
                }

                resultOffset += chunkSize;
            }

            return result;
        }

        private void ReadDirectoryEntries(byte[] bytes)
        {
            try
            {
                Entries = new List<CompoundDirectoryEntry>();
                using (var stream = new MemoryStream(bytes))
                {
                    using (var reader = new BinaryReader(stream))
                    {
                        RootEntry = ReadDirectoryEntry(reader);
                        Entries.Add(RootEntry);

                        while (stream.Position < stream.Length)
                        {
                            var entry = ReadDirectoryEntry(reader);
                            Entries.Add(entry);
                        }
                    }
                }
            }
            catch (EndOfStreamException ex)
            {
                throw new InvalidOperationException("The excel file may be corrupt or truncated. We've read past the end of the file.", ex);
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
        }

        private DateTime ReadFileTime(BinaryReader reader)
        {
            var d = reader.ReadInt64();
            if (d < 0 || d > DateTime.MaxValue.ToFileTime())
            {
                d = 0;
            }

            return DateTime.FromFileTime(d);
        }

        private CompoundHeader ReadHeader(BinaryReader reader)
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

            var chain = new List<uint>();
            for (int i = 0; i < 109; ++i)
            {
                chain.Add(reader.ReadUInt32());
            }

            result.First109DifSectorChain = chain;

            return result;
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
                    throw new InvalidOperationException("The excel file may be corrupt or truncated. We've read past the end of the file.", ex);
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
                throw new InvalidOperationException("The excel file may be corrupt or truncated. We've read past the end of the file.", ex);
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

        private void TrimSectorChain(List<uint> chain, FATMARKERS marker)
        {
            while (chain.Count > 0 && chain[chain.Count - 1] == (uint)marker)
            {
                chain.RemoveAt(chain.Count - 1);
            }
        }

        private uint GetNextSector(uint sector, List<uint> sectorTable)
        {
            if (sector < sectorTable.Count)
            {
                return sectorTable[(int)sector];
            }

            return (uint)FATMARKERS.FAT_EndOfChain;
        }
    }
}
