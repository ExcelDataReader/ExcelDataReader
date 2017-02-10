using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
	/// Represents formatted unicode string in SST
	/// </summary>
	internal class XlsFormattedUnicodeString : IXlsString
    {
		#region FormattedUnicodeStringFlags enum

		[Flags]
		public enum FormattedUnicodeStringFlags : byte
		{
			MultiByte = 0x01,
			HasExtendedString = 0x04,
			HasFormatting = 0x08,
		}

		#endregion

		private readonly byte[] m_bytes;
        private readonly uint m_offset;

        public XlsFormattedUnicodeString(byte[] bytes, uint offset)
        {
            m_bytes = bytes;
            m_offset = offset;
        }

		/// <summary>
		/// Count of characters in string
		/// </summary>
		public ushort CharacterCount => BitConverter.ToUInt16(m_bytes, (int)m_offset);

        /// <summary>
		/// String flags
		/// </summary>
		public FormattedUnicodeStringFlags Flags => (FormattedUnicodeStringFlags)Buffer.GetByte(m_bytes, (int)m_offset + 2);

        /// <summary>
		/// Checks if string has Extended record
		/// </summary>
		public bool HasExtString => (Flags & FormattedUnicodeStringFlags.HasExtendedString) == FormattedUnicodeStringFlags.HasExtendedString;

        /// <summary>
		/// Checks if string has formatting
		/// </summary>
		public bool HasFormatting => (Flags & FormattedUnicodeStringFlags.HasFormatting) == FormattedUnicodeStringFlags.HasFormatting;

        /// <summary>
		/// Checks if string is unicode
		/// </summary>
		public bool IsMultiByte => (Flags & FormattedUnicodeStringFlags.MultiByte) == FormattedUnicodeStringFlags.MultiByte;

        /// <summary>
		/// Returns number of formats used for formatting (0 if string has no formatting)
		/// </summary>
		public ushort FormatCount => HasFormatting ? BitConverter.ToUInt16(m_bytes, (int)m_offset + 3) : (ushort)0;

        /// <summary>
		/// Returns size of extended string in bytes, 0 if there is no one
		/// </summary>
		public uint ExtendedStringSize => HasExtString ? (uint)BitConverter.ToUInt16(m_bytes, (int)m_offset + ((HasFormatting) ? 5 : 3)) : 0;

        /// <summary>
		/// Returns head (before string data) size in bytes
		/// </summary>
		public uint HeadSize => (uint)(HasFormatting ? 2 : 0) + (uint)(HasExtString ? 4 : 0) + 3;

        /// <summary>
		/// Returns tail (after string data) size in bytes
		/// </summary>
		public uint TailSize => (uint)(HasFormatting ? 4 * FormatCount : 0) + (HasExtString ? ExtendedStringSize : 0);

        /// <summary>
		/// Returns size of whole record in bytes
		/// </summary>
		public uint Size
		{
			get
			{
				uint extraSize = (uint)(HasFormatting ? 2 + FormatCount * 4 : 0) +
								 (HasExtString ? 4 + ExtendedStringSize : 0) + 3;
				if (!IsMultiByte)
					return extraSize + CharacterCount;
				return extraSize + (uint)CharacterCount * 2;
			}
		}

		/// <summary>
		/// Returns string represented by this instance
		/// </summary>
		public string Value
		{
			get
			{
			    if (IsMultiByte)
			        return Encoding.Unicode.GetString(m_bytes, (int)(m_offset + HeadSize), CharacterCount * 2);

			    int len = CharacterCount;
			    int start = (int)HeadSize;
			    byte[] bytes = new byte[len * 2];
			    for (int i = 0; i < len; i++)
			    {
			        bytes[i * 2] = m_bytes[m_offset + start + i];
			    }

			    return Encoding.Unicode.GetString(bytes, 0, len * 2);
			}
		}
	}
}