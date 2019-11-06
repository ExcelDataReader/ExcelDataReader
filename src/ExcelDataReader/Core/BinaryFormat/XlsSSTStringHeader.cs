using System;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// [MS-XLS] 2.5.293 XLUnicodeRichExtendedString
    /// Word-sized formatted string in SST, stored as single or multibyte unicode characters potentially spanning multiple Continue records.
    /// </summary>
    internal class XlsSSTStringHeader
    {
        private readonly byte[] _bytes;
        private readonly int _offset;

        public XlsSSTStringHeader(byte[] bytes, int offset)
        {
            _bytes = bytes;
            _offset = offset;
        }

        [Flags]
        public enum FormattedUnicodeStringFlags : byte
        {
            MultiByte = 0x01,
            HasExtendedString = 0x04,
            HasFormatting = 0x08,
        }

        /// <summary>
        /// Gets the number of characters in the string.
        /// </summary>
        public ushort CharacterCount => BitConverter.ToUInt16(_bytes, _offset);

        /// <summary>
        /// Gets the flags.
        /// </summary>
        public FormattedUnicodeStringFlags Flags => (FormattedUnicodeStringFlags)Buffer.GetByte(_bytes, _offset + 2);

        /// <summary>
        /// Gets a value indicating whether the string has an extended record. 
        /// </summary>
        public bool HasExtString => (Flags & FormattedUnicodeStringFlags.HasExtendedString) == FormattedUnicodeStringFlags.HasExtendedString;

        /// <summary>
        /// Gets a value indicating whether the string has a formatting record.
        /// </summary>
        public bool HasFormatting => (Flags & FormattedUnicodeStringFlags.HasFormatting) == FormattedUnicodeStringFlags.HasFormatting;

        /// <summary>
        /// Gets a value indicating whether the string is a multibyte string or not.
        /// </summary>
        public bool IsMultiByte => (Flags & FormattedUnicodeStringFlags.MultiByte) == FormattedUnicodeStringFlags.MultiByte;

        /// <summary>
        /// Gets the number of formats used for formatting (0 if string has no formatting)
        /// </summary>
        public ushort FormatCount => HasFormatting ? BitConverter.ToUInt16(_bytes, (int)_offset + 3) : (ushort)0;

        /// <summary>
        /// Gets the size of extended string in bytes, 0 if there is no one
        /// </summary>
        public uint ExtendedStringSize => HasExtString ? (uint)BitConverter.ToUInt32(_bytes, (int)_offset + (HasFormatting ? 5 : 3)) : 0;

        /// <summary>
        /// Gets the head (before string data) size in bytes
        /// </summary>
        public uint HeadSize => (uint)(HasFormatting ? 2 : 0) + (uint)(HasExtString ? 4 : 0) + 3;

        /// <summary>
        /// Gets the tail (after string data) size in bytes
        /// </summary>
        public uint TailSize => (uint)(HasFormatting ? 4 * FormatCount : 0) + (HasExtString ? ExtendedStringSize : 0);
    }
}