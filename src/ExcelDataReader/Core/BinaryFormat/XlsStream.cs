using System;
using System.IO;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents an Excel file stream
    /// </summary>
    internal class XlsStream
    {
        private readonly XlsFat _fat;
        private readonly Stream _fileStream;
        private readonly XlsHeader _hdr;

        private readonly bool _isMini;
        private readonly XlsRootDirectory _rootDir;
        private readonly XlsFat _minifat;

        public XlsStream(XlsHeader hdr, uint startSector, bool isMini, XlsRootDirectory rootDir)
        {
            _fileStream = hdr.FileStream;
            _fat = hdr.Fat;
            _hdr = hdr;
            BaseSector = startSector;
            _isMini = isMini;
            _rootDir = rootDir;

            _minifat = _hdr.GetMiniFat();
        }

        /// <summary>
        /// Gets the offset of first stream sector
        /// </summary>
        public uint BaseOffset => (uint)((BaseSector + 1) * _hdr.SectorSize);

        /// <summary>
        /// Gets the number of first stream sector
        /// </summary>
        public uint BaseSector { get; }

        /// <summary>
        /// Reads stream data from file
        /// </summary>
        /// <returns>Stream data</returns>
        public byte[] ReadStream()
        {
            uint sector = BaseSector, prevSector = 0;
            int sectorSize = _isMini ? _hdr.MiniSectorSize : _hdr.SectorSize;
            var fat = _isMini ? _minifat : _fat;
            long offset = 0;
            if (_isMini && _rootDir != null)
            {
                offset = (_rootDir.RootEntry.StreamFirstSector + 1) * _hdr.SectorSize;
            }

            byte[] buff = new byte[sectorSize];
            byte[] ret;

            using (MemoryStream ms = new MemoryStream(sectorSize * 8))
            {
                lock (_fileStream)
                {
                    do
                    {
                        if (prevSector == 0 || (sector - prevSector) != 1)
                        {
                            var adjustedSector = _isMini ? sector : sector + 1; // standard sector is + 1 because header is first
                            _fileStream.Seek(adjustedSector * sectorSize + offset, SeekOrigin.Begin);
                        }

                        if (prevSector != 0 && prevSector == sector)
                            throw new InvalidOperationException("The excel file may be corrupt. We appear to be stuck");

                        prevSector = sector;
                        _fileStream.Read(buff, 0, sectorSize);
                        ms.Write(buff, 0, sectorSize);

                        sector = fat.GetNextSector(sector);

                        if (sector == 0)
                            throw new InvalidOperationException("Next sector cannot be 0. Possibly corrupt excel file");
                    }
                    while (sector != (uint)FATMARKERS.FAT_EndOfChain);
                }

                ret = ms.ToArray();
            }

            return ret;
        }
    }
}
