using System;
using System.Collections.Generic;
using System.IO;
using Excel;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Excel file header
    /// </summary>
    internal class XlsHeader
    {
        private readonly byte[] _bytes;

        private XlsFat _fat;
        private XlsFat _minifat;

        private XlsHeader(Stream file)
        {
            _bytes = new byte[512];
            FileStream = file;
        }

        /// <summary>
        /// Gets the file signature
        /// </summary>
        public ulong Signature => BitConverter.ToUInt64(_bytes, 0x0);

        /// <summary>
        /// Gets a value indicating whether the signature is valid. 
        /// </summary>
        public bool IsSignatureValid => Signature == 0xE11AB1A1E011CFD0;

        /// <summary>
        /// Gets the class id. Typically filled with zeroes
        /// </summary>
        public Guid ClassId
        {
            get
            {
                byte[] tmp = new byte[16];
                Buffer.BlockCopy(_bytes, 0x8, tmp, 0, 16);
                return new Guid(tmp);
            }
        }

        /// <summary>
        /// Gets the version. Must be 0x003E
        /// </summary>
        public ushort Version => BitConverter.ToUInt16(_bytes, 0x18);

        /// <summary>
        /// Gets the dll version. Must be 0x0003
        /// </summary>
        public ushort DllVersion => BitConverter.ToUInt16(_bytes, 0x1A);

        /// <summary>
        /// Gets the byte order. Must be 0xFFFE
        /// </summary>
        public ushort ByteOrder => BitConverter.ToUInt16(_bytes, 0x1C);

        /// <summary>
        /// Gets the sector size. Typically 512
        /// </summary>
        public int SectorSize => 1 << BitConverter.ToUInt16(_bytes, 0x1E);

        /// <summary>
        /// Gets the mini sector size. Typically 64
        /// </summary>
        public int MiniSectorSize => 1 << BitConverter.ToUInt16(_bytes, 0x20);

        /// <summary>
        /// Gets the number of FAT sectors
        /// </summary>
        public int FatSectorCount => BitConverter.ToInt32(_bytes, 0x2C);

        /// <summary>
        /// Gets the number of first Root Directory Entry (Property Set Storage, FAT Directory) sector
        /// </summary>
        public uint RootDirectoryEntryStart => BitConverter.ToUInt32(_bytes, 0x30);

        /// <summary>
        /// Gets the transaction signature, 0 for Excel
        /// </summary>
        public uint TransactionSignature => BitConverter.ToUInt32(_bytes, 0x34);

        /// <summary>
        /// Gets the maximum size for small stream, typically 4096 bytes
        /// </summary>
        public uint MiniStreamCutoff => BitConverter.ToUInt32(_bytes, 0x38);

        /// <summary>
        /// Gets the first sector of Mini FAT, FAT_EndOfChain if there's no one
        /// </summary>
        public uint MiniFatFirstSector => BitConverter.ToUInt32(_bytes, 0x3C);

        /// <summary>
        /// Gets the number of sectors in Mini FAT, 0 if there's no one
        /// </summary>
        public int MiniFatSectorCount => BitConverter.ToInt32(_bytes, 0x40);

        /// <summary>
        /// Gets the first sector of DIF, FAT_EndOfChain if there's no one
        /// </summary>
        public uint DifFirstSector => BitConverter.ToUInt32(_bytes, 0x44);

        /// <summary>
        /// Gets the number of sectors in DIF, 0 if there's no one
        /// </summary>
        public int DifSectorCount => BitConverter.ToInt32(_bytes, 0x48);

        public Stream FileStream { get; }

        /// <summary>
        /// Gets the full FAT table, including DIF sectors
        /// </summary>
        public XlsFat Fat
        {
            get
            {
                if (_fat != null)
                    return _fat;

                uint value;
                int sectorSize = SectorSize;
                List<uint> sectors = new List<uint>(FatSectorCount);
                for (int i = 0x4C; i < sectorSize; i += 4)
                {
                    value = BitConverter.ToUInt32(_bytes, i);
                    if (value == (uint)FATMARKERS.FAT_FreeSpace)
                        goto XlsHeader_Fat_Ready;
                    sectors.Add(value);
                }

                int difCount;
                if ((difCount = DifSectorCount) == 0)
                    goto XlsHeader_Fat_Ready;
                lock (FileStream)
                {
                    uint difSector = DifFirstSector;
                    byte[] buff = new byte[sectorSize];
                    uint prevSector = 0;
                    while (difCount > 0)
                    {
                        sectors.Capacity += 128;
                        if (prevSector == 0 || (difSector - prevSector) != 1)
                            FileStream.Seek((difSector + 1) * sectorSize, SeekOrigin.Begin);
                        prevSector = difSector;
                        FileStream.Read(buff, 0, sectorSize);
                        for (int i = 0; i < 508; i += 4)
                        {
                            value = BitConverter.ToUInt32(buff, i);
                            if (value == (uint)FATMARKERS.FAT_FreeSpace)
                                goto XlsHeader_Fat_Ready;
                            sectors.Add(value);
                        }

                        value = BitConverter.ToUInt32(buff, 508);
                        if (value == (uint)FATMARKERS.FAT_FreeSpace)
                            break;
                        if (difCount-- > 1)
                            difSector = value;
                        else
                            sectors.Add(value);
                    }
                }

                XlsHeader_Fat_Ready:
                _fat = new XlsFat(this, sectors);
                return _fat;
            }
        }

        /// <summary>
        /// Reads Excel header from Stream
        /// </summary>
        /// <param name="file">Stream with Excel file</param>
        /// <returns>XlsHeader representing specified file</returns>
        public static XlsHeader ReadHeader(Stream file)
        {
            XlsHeader hdr = new XlsHeader(file);
            lock (file)
            {
                file.Seek(0, SeekOrigin.Begin);
                file.Read(hdr._bytes, 0, 512);
            }

            if (!hdr.IsSignatureValid)
                throw new HeaderException(Errors.ErrorHeaderSignature);
            if (hdr.ByteOrder != 0xFFFE && hdr.ByteOrder != 0xFFFF) // Some broken xls files uses 0xFFFF
                throw new FormatException(Errors.ErrorHeaderOrder);
            return hdr;
        }

        /// <summary>
        /// Returns mini FAT table
        /// </summary>
        public XlsFat GetMiniFat()
        {
            if (_minifat != null)
                return _minifat;
            
            // if no minifat then return null
            if (MiniFatSectorCount == 0/* || MiniSectorSize == 0xFFFFFFFE*/)
                return null;

            List<uint> sectors = new List<uint>(MiniFatSectorCount);

            // find the sector where the minifat starts
            var miniFatStartSector = BitConverter.ToUInt32(_bytes, 0x3c);
            sectors.Add(miniFatStartSector);
            /*
            lock (m_file)
            {
                //work out the file location of minifat then read each sector
                var miniFatStartOffset = (miniFatStartSector + 1) * SectorSize;
                var miniFatSize = MiniFatSectorCount * SectorSize;
                m_file.Seek(miniFatStartOffset, SeekOrigin.Begin);

                byte[] sectorBuff = new byte[SectorSize];

                for (var i = 0; i < MiniFatSectorCount; i += SectorSize)
                {
                    m_file.Read(sectorBuff, 0, 4);
                    var secId = BitConverter.ToUInt32(sectorBuff, 0);
                    sectors.Add(secId);
                }
            }
            */
            _minifat = new XlsFat(this, sectors);
            return _minifat;
        }
    }
}
