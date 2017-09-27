using System;
using System.Collections.Generic;

namespace ExcelDataReader.Core.CompoundFormat
{
    /// <summary>
    /// Represents Excel file header
    /// </summary>
    internal class CompoundHeader
    {
        /// <summary>
        /// Gets or sets the file signature
        /// </summary>
        public ulong Signature { get; set; }

        /// <summary>
        /// Gets a value indicating whether the signature is valid. 
        /// </summary>
        public bool IsSignatureValid => Signature == 0xE11AB1A1E011CFD0;

        /// <summary>
        /// Gets or sets the class id. Typically filled with zeroes
        /// </summary>
        public Guid ClassId { get; set; }

        /// <summary>
        /// Gets or sets the version. Must be 0x003E
        /// </summary>
        public ushort Version { get; set; }

        /// <summary>
        /// Gets or sets the dll version. Must be 0x0003
        /// </summary>
        public ushort DllVersion { get; set; }

        /// <summary>
        /// Gets or sets the byte order. Must be 0xFFFE
        /// </summary>
        public ushort ByteOrder { get; set; }

        /// <summary>
        /// Gets or sets the sector size in Pot
        /// </summary>
        public int SectorSizeInPot { get; set; }

        /// <summary>
        /// Gets the sector size. Typically 512
        /// </summary>
        public int SectorSize => 1 << SectorSizeInPot;

        /// <summary>
        /// Gets or sets the mini sector size in Pot
        /// </summary>
        public int MiniSectorSizeInPot { get; set; }

        /// <summary>
        /// Gets the mini sector size. Typically 64
        /// </summary>
        public int MiniSectorSize => 1 << MiniSectorSizeInPot;

        /// <summary>
        /// Gets or sets the number of directory sectors. If Major Version is 3, the Number of 
        /// Directory Sectors MUST be zero. This field is not supported for version 3 compound files
        /// </summary>
        public int DirectorySectorCount { get; set; }

        /// <summary>
        /// Gets or sets the number of FAT sectors
        /// </summary>
        public int FatSectorCount { get; set; }

        /// <summary>
        /// Gets or sets the number of first Root Directory Entry (Property Set Storage, FAT Directory) sector
        /// </summary>
        public uint RootDirectoryEntryStart { get; set; }

        /// <summary>
        /// Gets or sets the transaction signature, 0 for Excel
        /// </summary>
        public uint TransactionSignature { get; set; }

        /// <summary>
        /// Gets or sets the maximum size for small stream, typically 4096 bytes
        /// </summary>
        public uint MiniStreamCutoff { get; set; }

        /// <summary>
        /// Gets or sets the first sector of Mini FAT, FAT_EndOfChain if there's no one
        /// </summary>
        public uint MiniFatFirstSector { get; set; }

        /// <summary>
        /// Gets or sets the number of sectors in Mini FAT, 0 if there's no one
        /// </summary>
        public int MiniFatSectorCount { get; set; }

        /// <summary>
        /// Gets or sets the first sector of DIF, FAT_EndOfChain if there's no one
        /// </summary>
        public uint DifFirstSector { get; set; }

        /// <summary>
        /// Gets or sets the number of sectors in DIF, 0 if there's no one
        /// </summary>
        public int DifSectorCount { get; set; }

        /// <summary>
        /// Gets or sets the first 109 locations in the DIF sector chain
        /// </summary>
        public List<uint> First109DifSectorChain { get; set; }
    }
}
