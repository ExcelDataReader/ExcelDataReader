using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat
{
    public class XlsDocument
    {
        public XlsDocument(Stream stream)
        {
            var reader = new BinaryReader(stream);

            Header = ReadHeader(reader);

            if (Header.IsRawBiffStream)
                throw new NotSupportedException("File appears to be a raw BIFF stream which isn't supported (BIFF" + Header.RawBiffVersion + ").");
            if (!Header.IsSignatureValid)
                throw new HeaderException(Errors.ErrorHeaderSignature);
            if (Header.ByteOrder != 0xFFFE && Header.ByteOrder != 0xFFFF) // Some broken xls files uses 0xFFFF
                throw new FormatException(Errors.ErrorHeaderOrder);

            var difSectorChain = ReadDifSectorChain(reader);
            SectorTable = ReadSectorTable(reader, difSectorChain);

            var miniChain = GetSectorChain(Header.MiniFatFirstSector, SectorTable);
            MiniSectorTable = ReadSectorTable(reader, miniChain);

            var bytes = ReadStream(stream, Header.RootDirectoryEntryStart, false);
            ReadDirectoryEntries(bytes);
        }

        internal XlsHeader Header { get; }

        internal List<uint> SectorTable { get; }

        internal List<uint> MiniSectorTable { get; }

        internal XlsDirectoryEntry RootEntry { get; set; }

        internal List<XlsDirectoryEntry> Entries { get; set; }

        internal static bool CheckRawBiffStream(byte[] bytes, out int version)
        {
            if (bytes.Length < 8)
            {
                throw new ArgumentException("Needs at least 8 bytes to probe", nameof(bytes));
            }

            ushort rid = BitConverter.ToUInt16(bytes, 0);
            ushort size = BitConverter.ToUInt16(bytes, 2);
            version = BitConverter.ToUInt16(bytes, 4);
            ushort type = BitConverter.ToUInt16(bytes, 6);

            switch (rid)
            {
                case 0x0009: // BIFF2
                    if (size != 4)
                        return false;
                    if (type != 0x10 && type != 0x20 && type != 0x40)
                        return false;
                    return true;
                case 0x0209: // BIFF3
                    if (size != 6)
                        return false;
                    if (type != 0x10 && type != 0x20 && type != 0x40 && type != 0x0100)
                        return false;
                    /* removed this additional check to keep the probe at 8 bytes
                    ushort notUsed = BitConverter.ToUInt16(bytes, 8);
                    if (notUsed != 0x00)
                        return false;*/
                    return true;
                case 0x0809: // BIFF5/BIFF8
                    if (size != 8 || size != 16)
                        return false;
                    if (version != 0x0500 && version != 0x600)
                        return false;
                    if (type != 0x5 && type != 0x6 && type != 0x10 && type != 0x20 && type != 0x40 && type != 0x0100)
                        return false;
                    /* removed this additional check to keep the probe at 8 bytes
                    ushort identifier = BitConverter.ToUInt16(bytes, 10);
                    if (identifier == 0)
                        return false;*/
                    return true;
            }

            return false;
        }

        internal XlsDirectoryEntry FindEntry(string entryName)
        {
            foreach (var e in Entries)
            {
                if (string.Equals(e.EntryName, entryName, StringComparison.CurrentCultureIgnoreCase))
                    return e;
            }

            return null;
        }

        /// <summary>
        /// Reads bytes from a regular or mini stream.
        /// </summary>
        internal byte[] ReadStream(Stream stream, uint baseSector, bool isMini)
        {
            if (isMini)
            {
                return ReadMiniStream(stream, baseSector);
            }
            else
            {
                return ReadRegularStream(stream, baseSector);
            }
        }

        /// <summary>
        /// Reads bytes from the mini stream stored inside the Root Entry's stream.
        /// </summary>
        private byte[] ReadMiniStream(Stream stream, uint baseSector)
        {
            var chain = GetSectorChain(baseSector, MiniSectorTable);
            var rootStreamChain = GetSectorChain(RootEntry.StreamFirstSector, SectorTable);

            var result = new byte[Header.MiniSectorSize * chain.Count];
            int resultOffset = 0;
            foreach (var sector in chain)
            {
                // Convert to sector+offset in the root stream
                var miniStreamOffset = (int)GetMiniSectorOffset(sector);

                var rootSector = rootStreamChain[miniStreamOffset / Header.SectorSize];
                var rootOffset = miniStreamOffset % Header.SectorSize;

                stream.Seek(GetSectorOffset(rootSector) + rootOffset, SeekOrigin.Begin);
                stream.Read(result, resultOffset, Header.MiniSectorSize);
                resultOffset += Header.MiniSectorSize;
            }

            return result;
        }

        private byte[] ReadRegularStream(Stream stream, uint baseSector)
        {
            var sectorSize = Header.SectorSize;
            var chain = GetSectorChain(baseSector, SectorTable);

            var result = new byte[sectorSize * chain.Count];
            int resultOffset = 0;
            foreach (var sector in chain)
            {
                stream.Seek(GetSectorOffset(sector), SeekOrigin.Begin);
                stream.Read(result, resultOffset, sectorSize);
                resultOffset += sectorSize;
            }

            return result;
        }

        private int ReadSector(Stream stream, uint sector, byte[] result, int offset)
        {
            stream.Seek(GetSectorOffset(sector), SeekOrigin.Begin);
            return stream.Read(result, offset, Header.SectorSize);
        }

        private void ReadDirectoryEntries(byte[] bytes)
        {
            try
            {
                Entries = new List<XlsDirectoryEntry>();
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

        private XlsDirectoryEntry ReadDirectoryEntry(BinaryReader reader)
        {
            var result = new XlsDirectoryEntry();
            var name = reader.ReadBytes(64);
            var nameLength = reader.ReadUInt16();

            if (nameLength > 0)
            {
                result.EntryName = Encoding.Unicode.GetString(name, 0, nameLength).TrimEnd('\0');
            }

            result.EntryType = (STGTY)reader.ReadByte();
            result.EntryColor = (DECOLOR)reader.ReadByte();
            result.LeftSiblingSid = reader.ReadUInt32();
            result.RightSiblingSid = reader.ReadUInt32();
            result.ChildSid = reader.ReadUInt32();
            result.ClassId = new Guid(reader.ReadBytes(16));
            result.UserFlags = reader.ReadUInt32();
            result.CreationTime = DateTime.FromFileTime(reader.ReadInt64());
            result.LastWriteTime = DateTime.FromFileTime(reader.ReadInt64());
            result.StreamFirstSector = reader.ReadUInt32();
            result.StreamSize = reader.ReadUInt32();
            result.PropType = reader.ReadUInt32();
            result.IsEntryMiniStream = result.StreamSize < Header.MiniStreamCutoff;
            return result;
        }

        private XlsHeader ReadHeader(BinaryReader reader)
        {
            var result = new XlsHeader();
            var signature = reader.ReadBytes(8);

            if (CheckRawBiffStream(signature, out int version))
            {
                result.IsRawBiffStream = true;
                result.RawBiffVersion = version;
                return result;
            }

            result.Signature = BitConverter.ToUInt64(signature, 0);
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
                var difBytes = new byte[Header.DifSectorCount * Header.SectorSize];

                for (var i = 0; i < Header.DifSectorCount; ++i)
                {
                    var difSector = (uint)(Header.DifFirstSector + i);
                    var difContent = ReadSectorAsUInt32s(reader, difSector);
                    difSectorChain.AddRange(difContent);
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

        private long GetMiniSectorOffset(uint sector)
        {
            return Header.MiniSectorSize * sector;
        }

        private long GetSectorOffset(uint sector)
        {
            return 512 + Header.SectorSize * sector;
        }

        private List<uint> GetSectorChain(uint sector, List<uint> sectorTable)
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

        private uint GetNextSector(uint sector, List<uint> sectorTable)
        {
            if (sector < sectorTable.Count)
            {
                return sectorTable[(int)sector];
            }
            else
            {
                return (uint)FATMARKERS.FAT_EndOfChain;
            }
        }
    }
}
