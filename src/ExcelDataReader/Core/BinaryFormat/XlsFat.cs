using System;
using System.Collections.Generic;
using System.IO;
using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Excel file FAT table
    /// </summary>
    internal class XlsFat
    {
        private readonly List<uint> _fat;

        /// <summary>
        /// Initializes a new instance of the <see cref="XlsFat"/> class.
        /// </summary>
        /// <remarks>
        /// Constructs FAT table from list of sectors.
        /// </remarks>
        /// <param name="hdr">XlsHeader</param>
        /// <param name="sectors">Sectors list</param>
        public XlsFat(XlsHeader hdr, List<uint> sectors)
        {
            Header = hdr;
            SectorsForFat = sectors.Count;
            int sizeOfSector = hdr.SectorSize;
            uint prevSector = 0;

            // calc offset of stream . If mini stream then find mini stream container stream
            /*
            long offset = 0;
            if (rootDir != null)
                offset = isMini ? (hdr.MiniFatFirstSector + 1) * hdr.SectorSize : 0;
            */

            byte[] buff = new byte[sizeOfSector];
            Stream file = hdr.FileStream;
            using (MemoryStream ms = new MemoryStream(sizeOfSector * SectorsForFat))
            {
                lock (file)
                {
                    for (int i = 0; i < sectors.Count; i++)
                    {
                        uint sector = sectors[i];
                        if (prevSector == 0 || sector - prevSector != 1)
                            file.Seek((sector + 1) * sizeOfSector, SeekOrigin.Begin);
                        prevSector = sector;
                        file.Read(buff, 0, sizeOfSector);
                        ms.Write(buff, 0, sizeOfSector);
                    }
                }

                ms.Seek(0, SeekOrigin.Begin);
                using (BinaryReader rd = new BinaryReader(ms))
                {
                    SectorsCount = (int)ms.Length / 4;
                    _fat = new List<uint>(SectorsCount);
                    for (int i = 0; i < SectorsCount; i++)
                        _fat.Add(rd.ReadUInt32());
                }
            }
        }

        /// <summary>
        /// Gets the number of sectors used by FAT itself
        /// </summary>
        public int SectorsForFat { get; }

        /// <summary>
        /// Gets the number of sectors described by FAT
        /// </summary>
        public int SectorsCount { get; }

        /// <summary>
        /// Gets the underlying XlsHeader object
        /// </summary>
        public XlsHeader Header { get; }

        /// <summary>
        /// Returns next data sector using FAT
        /// </summary>
        /// <param name="sector">Current data sector</param>
        /// <returns>Next data sector</returns>
        public uint GetNextSector(uint sector)
        {
            if (_fat.Count <= sector)
                throw new ArgumentException(Errors.ErrorFatBadSector);
            uint value = _fat[(int)sector];
            if (value == (uint)FATMARKERS.FAT_FatSector || value == (uint)FATMARKERS.FAT_DifSector)
                throw new InvalidOperationException(Errors.ErrorFatRead);
            return value;
        }
    }
}
