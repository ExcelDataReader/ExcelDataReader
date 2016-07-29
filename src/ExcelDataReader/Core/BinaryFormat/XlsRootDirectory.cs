using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using ExcelDataReader.Misc;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents Root Directory in file
	/// </summary>
	internal class XlsRootDirectory
	{
		private readonly List<XlsDirectoryEntry> m_entries;
		private readonly XlsDirectoryEntry m_root;

		/// <summary>
		/// Creates Root Directory catalog from XlsHeader
		/// </summary>
		/// <param name="hdr">XlsHeader object</param>
		public XlsRootDirectory(XlsHeader hdr)
		{
			XlsStream stream = new XlsStream(hdr, hdr.RootDirectoryEntryStart, false, null);
			byte[] array = stream.ReadStream();
			byte[] tmp;
			XlsDirectoryEntry entry;
			List<XlsDirectoryEntry> entries = new List<XlsDirectoryEntry>();
			for (int i = 0; i < array.Length; i += XlsDirectoryEntry.Length)
			{
				tmp = new byte[XlsDirectoryEntry.Length];
				Buffer.BlockCopy(array, i, tmp, 0, tmp.Length);
				entries.Add(new XlsDirectoryEntry(tmp, hdr));
			}
			m_entries = entries;
			for (int i = 0; i < entries.Count; i++)
			{
				entry = entries[i];

				//Console.WriteLine("Directory Entry:{0} type:{1}, firstsector:{2}, streamSize:{3}, isEntryMiniStream:{4}", entry.EntryName, entry.EntryType.ToString(), entry.StreamFirstSector, entry.StreamSize, entry.IsEntryMiniStream);
				if (m_root == null && entry.EntryType == STGTY.STGTY_ROOT)
					m_root = entry;
				if (entry.ChildSid != (uint)FATMARKERS.FAT_FreeSpace)
					entry.Child = entries[(int)entry.ChildSid];
				if (entry.LeftSiblingSid != (uint)FATMARKERS.FAT_FreeSpace)
					entry.LeftSibling = entries[(int)entry.LeftSiblingSid];
				if (entry.RightSiblingSid != (uint)FATMARKERS.FAT_FreeSpace)
					entry.RightSibling = entries[(int)entry.RightSiblingSid];
			}
			stream.CalculateMiniFat(this);
		}

		/// <summary>
		/// Returns all entries in Root Directory
		/// </summary>
		public ReadOnlyCollection<XlsDirectoryEntry> Entries
		{
			get { return new ReadOnlyCollection<XlsDirectoryEntry>(m_entries); }
		}

		/// <summary>
		/// Returns the Root Entry
		/// </summary>
		public XlsDirectoryEntry RootEntry
		{
			get { return m_root; }
		}

		/// <summary>
		/// Searches for first matching entry by its name
		/// </summary>
		/// <param name="EntryName">String name of entry</param>
		/// <returns>Entry if found, null otherwise</returns>
		public XlsDirectoryEntry FindEntry(string EntryName)
		{
			foreach (XlsDirectoryEntry e in m_entries)
			{
                if (string.Equals(e.EntryName, EntryName, StringComparison.CurrentCultureIgnoreCase))
					return e;
			}
			return null;
		}
	}
}
