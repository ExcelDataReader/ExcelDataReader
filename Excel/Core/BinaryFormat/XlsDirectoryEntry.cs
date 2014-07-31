using System;
using System.IO;
using System.Text;

namespace Excel.Core.BinaryFormat
{
	/// <summary>
	/// Represents single Root Directory record
	/// </summary>
	internal class XlsDirectoryEntry
	{
		/// <summary>
		/// Length of structure in bytes
		/// </summary>
		public const int Length = 0x80;

		private readonly byte[] m_bytes;
		private XlsDirectoryEntry m_child;
		private XlsDirectoryEntry m_leftSibling;
		private XlsDirectoryEntry m_rightSibling;
		private XlsHeader m_header;

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="bytes">byte array representing current object</param>
		/// <param name="header"></param>
		public XlsDirectoryEntry(byte[] bytes, XlsHeader header)
		{
			if (bytes.Length < Length)
				throw new Excel.Exceptions.BiffRecordException(Errors.ErrorDirectoryEntryArray);
			m_bytes = bytes;
			m_header = header;
		}

		/// <summary>
		/// Returns name of directory entry
		/// </summary>
		public string EntryName
		{
			get { return Encoding.Unicode.GetString(m_bytes, 0x0, EntryLength).TrimEnd('\0'); }
		}

		/// <summary>
		/// Returns size in bytes of entry's name (with terminating zero)
		/// </summary>
		public ushort EntryLength
		{
			get { return BitConverter.ToUInt16(m_bytes, 0x40); }
		}

		/// <summary>
		/// Returns entry type
		/// </summary>
		public STGTY EntryType
		{
			get { return (STGTY)Buffer.GetByte(m_bytes, 0x42); }
		}

		/// <summary>
		/// Retuns entry "color" in directory tree
		/// </summary>
		public DECOLOR EntryColor
		{
			get { return (DECOLOR)Buffer.GetByte(m_bytes, 0x43); }
		}

		/// <summary>
		/// Returns SID of left sibling
		/// </summary>
		/// <remarks>0xFFFFFFFF if there's no one</remarks>
		public uint LeftSiblingSid
		{
			get { return BitConverter.ToUInt32(m_bytes, 0x44); }
		}

		/// <summary>
		/// Returns left sibling
		/// </summary>
		public XlsDirectoryEntry LeftSibling
		{
			get { return m_leftSibling; }
			set { if (m_leftSibling == null) m_leftSibling = value; }
		}

		/// <summary>
		/// Returns SID of right sibling
		/// </summary>
		/// <remarks>0xFFFFFFFF if there's no one</remarks>
		public uint RightSiblingSid
		{
			get { return BitConverter.ToUInt32(m_bytes, 0x48); }
		}

		/// <summary>
		/// Returns right sibling
		/// </summary>
		public XlsDirectoryEntry RightSibling
		{
			get { return m_rightSibling; }
			set { if (m_rightSibling == null) m_rightSibling = value; }
		}

		/// <summary>
		/// Returns SID of first child (if EntryType is STGTY_STORAGE)
		/// </summary>
		/// <remarks>0xFFFFFFFF if there's no one</remarks>
		public uint ChildSid
		{
			get { return BitConverter.ToUInt32(m_bytes, 0x4C); }
		}

		/// <summary>
		/// Returns child
		/// </summary>
		public XlsDirectoryEntry Child
		{
			get { return m_child; }
			set { if (m_child == null) m_child = value; }
		}

		/// <summary>
		/// CLSID of container (if EntryType is STGTY_STORAGE)
		/// </summary>
		public Guid ClassId
		{
			get
			{
				byte[] tmp = new byte[16];
				Buffer.BlockCopy(m_bytes, 0x50, tmp, 0, 16);
				return new Guid(tmp);
			}
		}

		/// <summary>
		/// Returns user flags of container (if EntryType is STGTY_STORAGE)
		/// </summary>
		public uint UserFlags
		{
			get { return BitConverter.ToUInt32(m_bytes, 0x60); }
		}

		/// <summary>
		/// Returns creation time of entry
		/// </summary>
		public DateTime CreationTime
		{
			get { return DateTime.FromFileTime(BitConverter.ToInt64(m_bytes, 0x64)); }
		}

		/// <summary>
		/// Returns last modification time of entry
		/// </summary>
		public DateTime LastWriteTime
		{
			get { return DateTime.FromFileTime(BitConverter.ToInt64(m_bytes, 0x6C)); }
		}

		/// <summary>
		/// First sector of data stream (if EntryType is STGTY_STREAM)
		/// </summary>
		/// <remarks>if EntryType is STGTY_ROOT, this can be first sector of MiniStream</remarks>
		public uint StreamFirstSector
		{
			get { return BitConverter.ToUInt32(m_bytes, 0x74); }
		}

		/// <summary>
		/// Size of data stream (if EntryType is STGTY_STREAM)
		/// </summary>
		/// <remarks>if EntryType is STGTY_ROOT, this can be size of MiniStream</remarks>
		public uint StreamSize
		{
			get { return BitConverter.ToUInt32(m_bytes, 0x78); }
		}

		/// <summary>
		/// Determines whether this entry relats to a ministream
		/// </summary>
		public bool IsEntryMiniStream
		{
			get { return (StreamSize < m_header.MiniStreamCutoff); }
		}

		/// <summary>
		/// Reserved, must be 0
		/// </summary>
		public uint PropType
		{
			get { return BitConverter.ToUInt32(m_bytes, 0x7C); }
		}
	}
}
