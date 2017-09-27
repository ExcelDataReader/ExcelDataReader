using System;

namespace ExcelDataReader.Core.CompoundFormat
{
    /// <summary>
    /// Represents single Root Directory record
    /// </summary>
    internal class CompoundDirectoryEntry
    {
        /// <summary>
        /// Gets or sets the name of directory entry
        /// </summary>
        public string EntryName { get; set; }

        /// <summary>
        /// Gets or sets the entry type
        /// </summary>
        public STGTY EntryType { get; set; }

        /// <summary>
        /// Gets or sets the entry "color" in directory tree
        /// </summary>
        public DECOLOR EntryColor { get; set; }

        /// <summary>
        /// Gets or sets the SID of left sibling
        /// </summary>
        /// <remarks>0xFFFFFFFF if there's no one</remarks>
        public uint LeftSiblingSid { get; set; }

        /// <summary>
        /// Gets or sets the SID of right sibling
        /// </summary>
        /// <remarks>0xFFFFFFFF if there's no one</remarks>
        public uint RightSiblingSid { get; set; }

        /// <summary>
        /// Gets or sets the SID of first child (if EntryType is STGTY_STORAGE)
        /// </summary>
        /// <remarks>0xFFFFFFFF if there's no one</remarks>
        public uint ChildSid { get; set; }

        /// <summary>
        /// Gets or sets the CLSID of container (if EntryType is STGTY_STORAGE)
        /// </summary>
        public Guid ClassId { get; set; }

        /// <summary>
        /// Gets or sets the user flags of container (if EntryType is STGTY_STORAGE)
        /// </summary>
        public uint UserFlags { get; set; }

        /// <summary>
        /// Gets or sets the creation time of entry
        /// </summary>
        public DateTime CreationTime { get; set; }

        /// <summary>
        /// Gets or sets the last modification time of entry
        /// </summary>
        public DateTime LastWriteTime { get; set; }

        /// <summary>
        /// Gets or sets the first sector of data stream (if EntryType is STGTY_STREAM)
        /// </summary>
        /// <remarks>if EntryType is STGTY_ROOT, this can be first sector of MiniStream</remarks>
        public uint StreamFirstSector { get; set; }

        /// <summary>
        /// Gets or sets the size of data stream (if EntryType is STGTY_STREAM)
        /// </summary>
        /// <remarks>if EntryType is STGTY_ROOT, this can be size of MiniStream</remarks>
        public uint StreamSize { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this entry relats to a ministream
        /// </summary>
        public bool IsEntryMiniStream { get; set; }

        /// <summary>
        /// Gets or sets the prop type. Reserved, must be 0.
        /// </summary>
        public uint PropType { get; set; }
    }
}
