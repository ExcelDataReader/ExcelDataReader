using System;
using System.Text;
using Excel;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat
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

        private readonly byte[] _bytes;
        private readonly XlsHeader _header;
        private XlsDirectoryEntry _child;
        private XlsDirectoryEntry _leftSibling;
        private XlsDirectoryEntry _rightSibling;

        /// <summary>
        /// Initializes a new instance of the <see cref="XlsDirectoryEntry"/> class.
        /// </summary>
        /// <param name="bytes">byte array representing current object</param>
        /// <param name="header">The header.</param>
        public XlsDirectoryEntry(byte[] bytes, XlsHeader header)
        {
            if (bytes.Length < Length)
                throw new BiffRecordException(Errors.ErrorDirectoryEntryArray);
            _bytes = bytes;
            _header = header;
        }

        /// <summary>
        /// Gets the name of directory entry
        /// </summary>
        public string EntryName => Encoding.Unicode.GetString(_bytes, 0x0, EntryLength).TrimEnd('\0');

        /// <summary>
        /// Gets the size in bytes of entry's name (with terminating zero)
        /// </summary>
        public ushort EntryLength => BitConverter.ToUInt16(_bytes, 0x40);

        /// <summary>
        /// Gets the entry type
        /// </summary>
        public STGTY EntryType => (STGTY)Buffer.GetByte(_bytes, 0x42);

        /// <summary>
        /// Gets the entry "color" in directory tree
        /// </summary>
        public DECOLOR EntryColor => (DECOLOR)Buffer.GetByte(_bytes, 0x43);

        /// <summary>
        /// Gets the SID of left sibling
        /// </summary>
        /// <remarks>0xFFFFFFFF if there's no one</remarks>
        public uint LeftSiblingSid => BitConverter.ToUInt32(_bytes, 0x44);
        
        /// <summary>
        /// Gets or sets  the left sibling
        /// </summary>
        public XlsDirectoryEntry LeftSibling
        {
            get => _leftSibling;
            set { if (_leftSibling == null) _leftSibling = value; }
        }

        /// <summary>
        /// Gets the SID of right sibling
        /// </summary>
        /// <remarks>0xFFFFFFFF if there's no one</remarks>
        public uint RightSiblingSid => BitConverter.ToUInt32(_bytes, 0x48);
        
        /// <summary>
        /// Gets or sets the right sibling
        /// </summary>
        public XlsDirectoryEntry RightSibling
        {
            get => _rightSibling;
            set { if (_rightSibling == null) _rightSibling = value; }
        }

        /// <summary>
        /// Gets the SID of first child (if EntryType is STGTY_STORAGE)
        /// </summary>
        /// <remarks>0xFFFFFFFF if there's no one</remarks>
        public uint ChildSid => BitConverter.ToUInt32(_bytes, 0x4C);
        
        /// <summary>
        /// Gets or sets the child
        /// </summary>
        public XlsDirectoryEntry Child
        {
            get => _child;
            set { if (_child == null) _child = value; }
        }

        /// <summary>
        /// Gets the CLSID of container (if EntryType is STGTY_STORAGE)
        /// </summary>
        public Guid ClassId
        {
            get
            {
                byte[] tmp = new byte[16];
                Buffer.BlockCopy(_bytes, 0x50, tmp, 0, 16);
                return new Guid(tmp);
            }
        }

        /// <summary>
        /// Gets the user flags of container (if EntryType is STGTY_STORAGE)
        /// </summary>
        public uint UserFlags => BitConverter.ToUInt32(_bytes, 0x60);

        /// <summary>
        /// Gets the creation time of entry
        /// </summary>
        public DateTime CreationTime => DateTime.FromFileTime(BitConverter.ToInt64(_bytes, 0x64));

        /// <summary>
        /// Gets the last modification time of entry
        /// </summary>
        public DateTime LastWriteTime => DateTime.FromFileTime(BitConverter.ToInt64(_bytes, 0x6C));

        /// <summary>
        /// Gets the first sector of data stream (if EntryType is STGTY_STREAM)
        /// </summary>
        /// <remarks>if EntryType is STGTY_ROOT, this can be first sector of MiniStream</remarks>
        public uint StreamFirstSector => BitConverter.ToUInt32(_bytes, 0x74);

        /// <summary>
        /// Gets the size of data stream (if EntryType is STGTY_STREAM)
        /// </summary>
        /// <remarks>if EntryType is STGTY_ROOT, this can be size of MiniStream</remarks>
        public uint StreamSize => BitConverter.ToUInt32(_bytes, 0x78);

        /// <summary>
        /// Gets a value indicating whether this entry relats to a ministream
        /// </summary>
        public bool IsEntryMiniStream => StreamSize < _header.MiniStreamCutoff;

        /// <summary>
        /// Gets the prop type. Reserved, must be 0.
        /// </summary>
        public uint PropType => BitConverter.ToUInt32(_bytes, 0x7C);
    }
}
