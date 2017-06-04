using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Root Directory in file
    /// </summary>
    internal class XlsRootDirectory
    {
        private readonly List<XlsDirectoryEntry> _entries;

        /// <summary>
        /// Initializes a new instance of the <see cref="XlsRootDirectory"/> class.
        /// </summary>
        /// <param name="hdr">The header.</param>
        public XlsRootDirectory(XlsHeader hdr)
        {
            XlsStream stream = new XlsStream(hdr, hdr.RootDirectoryEntryStart, false, null);
            byte[] array = stream.ReadStream();
            List<XlsDirectoryEntry> entries = new List<XlsDirectoryEntry>();
            for (int i = 0; i < array.Length; i += XlsDirectoryEntry.Length)
            {
                byte[] tmp = new byte[XlsDirectoryEntry.Length];
                Buffer.BlockCopy(array, i, tmp, 0, tmp.Length);
                entries.Add(new XlsDirectoryEntry(tmp, hdr));
            }

            _entries = entries;
            for (int i = 0; i < entries.Count; i++)
            {
                XlsDirectoryEntry entry = entries[i];

                // Console.WriteLine("Directory Entry:{0} type:{1}, firstsector:{2}, streamSize:{3}, isEntryMiniStream:{4}", entry.EntryName, entry.EntryType.ToString(), entry.StreamFirstSector, entry.StreamSize, entry.IsEntryMiniStream);
                if (RootEntry == null && entry.EntryType == STGTY.STGTY_ROOT)
                    RootEntry = entry;
                if (entry.ChildSid != (uint)FATMARKERS.FAT_FreeSpace)
                    entry.Child = entries[(int)entry.ChildSid];
                if (entry.LeftSiblingSid != (uint)FATMARKERS.FAT_FreeSpace)
                    entry.LeftSibling = entries[(int)entry.LeftSiblingSid];
                if (entry.RightSiblingSid != (uint)FATMARKERS.FAT_FreeSpace)
                    entry.RightSibling = entries[(int)entry.RightSiblingSid];
            }
        }

        /// <summary>
        /// Gets all entries in Root Directory
        /// </summary>
        public ReadOnlyCollection<XlsDirectoryEntry> Entries => new ReadOnlyCollection<XlsDirectoryEntry>(_entries);

        /// <summary>
        /// Gets the Root Entry
        /// </summary>
        public XlsDirectoryEntry RootEntry { get; }

        /// <summary>
        /// Searches for first matching entry by its name
        /// </summary>
        /// <param name="entryName">String name of entry</param>
        /// <returns>Entry if found, null otherwise</returns>
        public XlsDirectoryEntry FindEntry(string entryName)
        {
            foreach (XlsDirectoryEntry e in _entries)
            {
                if (string.Equals(e.EntryName, entryName, StringComparison.CurrentCultureIgnoreCase))
                    return e;
            }

            return null;
        }
    }
}
