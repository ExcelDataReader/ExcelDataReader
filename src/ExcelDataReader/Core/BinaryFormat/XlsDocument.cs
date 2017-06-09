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

            SectorTable = ReadSectorTable(reader, Header.DifSectorChain);

            var miniChain = GetSectorChain(Header.MiniFatFirstSector, false);
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

        internal byte[] ReadStream(Stream stream, uint baseSector, bool isMini)
        {
            int sectorSize;
            List<uint> rootStreamChain;

            if (isMini)
            {
                sectorSize = Header.MiniSectorSize;
                rootStreamChain = GetSectorChain(RootEntry.StreamFirstSector, false);
            }
            else
            {
                sectorSize = Header.SectorSize;
                rootStreamChain = null;
            }

            var chain = GetSectorChain(baseSector, isMini);
            var result = new byte[sectorSize * chain.Count];
            int resultOffset = 0;
            foreach (var sector in chain)
            {
                if (isMini)
                {
                    // Convert to sector+offset in the root stream
                    var miniStreamOffset = (int)GetSectorOffset(sector, isMini);

                    var rootSector = rootStreamChain[miniStreamOffset / Header.SectorSize];
                    var rootOffset = miniStreamOffset % Header.SectorSize;

                    stream.Seek(GetSectorOffset(rootSector, false) + rootOffset, SeekOrigin.Begin);

                    /*
                    // Previous (but incorrect) behavior: assume root stream is continous in the file
                    var miniOffset = GetSectorOffset(RootEntry.StreamFirstSector, false);
                    stream.Seek(miniOffset + GetSectorOffset(sector, isMini), SeekOrigin.Begin);
                    */
                }
                else
                {
                    stream.Seek(GetSectorOffset(sector, isMini), SeekOrigin.Begin);
                }

                stream.Read(result, resultOffset, sectorSize);
                resultOffset += sectorSize;
            }

            return result;
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
            reader.ReadBytes(10); // skip 10 unused bytes
            result.FatSectorCount = reader.ReadInt32();
            result.RootDirectoryEntryStart = reader.ReadUInt32();
            reader.ReadBytes(4); // skip 4 unused bytes
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

            while (chain.Count > 0 && chain[chain.Count - 1] == (uint)FATMARKERS.FAT_FreeSpace)
            {
                chain.RemoveAt(chain.Count - 1);
            }

            result.DifSectorChain = chain;

            return result;
        }

        private List<uint> GetSectorChain(uint sector, bool mini)
        {
            List<uint> chain = new List<uint>();
            while (sector != (uint)FATMARKERS.FAT_EndOfChain)
            {
                if (sector != (uint)FATMARKERS.FAT_FreeSpace)
                {
                    chain.Add(sector);
                }

                sector = GetNextSector(sector, mini);
            }

            return chain;
        }

        private long GetSectorOffset(uint sector, bool mini)
        {
            if (mini)
            {
                return Header.MiniSectorSize * sector;
            }
            else
            {
                return 512 + Header.SectorSize * sector;
            }
        }

        private uint GetNextSector(uint sector, bool mini)
        {
            if (mini)
            {
                if (sector < MiniSectorTable.Count)
                {
                    return MiniSectorTable[(int)sector];
                }
                else
                {
                    return (uint)FATMARKERS.FAT_EndOfChain;
                }
            }
            else
            {
                if (sector < SectorTable.Count)
                {
                    return SectorTable[(int)sector];
                }
                else
                {
                    return (uint)FATMARKERS.FAT_EndOfChain;
                }
            }
        }

        private List<uint> ReadSectorTable(BinaryReader reader, List<uint> chain)
        {
            var sectorTable = new List<uint>(Header.SectorSize / 4);
            try
            {
                foreach (var sector in chain)
                {
                    reader.BaseStream.Seek(GetSectorOffset(sector, false), SeekOrigin.Begin);
                    for (var i = 0; i < Header.SectorSize / 4; ++i)
                    {
                        var sectorTableSector = reader.ReadUInt32();
                        sectorTable.Add(sectorTableSector);
                    }
                }
            }
            catch (EndOfStreamException ex)
            {
                throw new InvalidOperationException("The excel file may be corrupt or truncated. We've read past the end of the file.", ex);
            }

            while (sectorTable.Count > 0 && sectorTable[sectorTable.Count - 1] == (uint)FATMARKERS.FAT_FreeSpace)
            {
                sectorTable.RemoveAt(sectorTable.Count - 1);
            }

            return sectorTable;
        }

    }
}
