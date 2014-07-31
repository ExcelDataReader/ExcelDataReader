using System;
using System.Text;

namespace Excel.Core.BinaryFormat
{
	/// <summary>
	/// Represents formatted unicode string in SST
	/// </summary>
	internal class XlsFormattedUnicodeString
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

		protected byte[] m_bytes;
		protected uint m_offset;

		public XlsFormattedUnicodeString(byte[] bytes, uint offset)
		{
			m_bytes = bytes;
			m_offset = offset;
		}

		/// <summary>
		/// Count of characters in string
		/// </summary>
		public ushort CharacterCount
		{
			get { return BitConverter.ToUInt16(m_bytes, (int)m_offset); }
		}

		/// <summary>
		/// String flags
		/// </summary>
		public FormattedUnicodeStringFlags Flags
		{
			get { return (FormattedUnicodeStringFlags)Buffer.GetByte(m_bytes, (int)m_offset + 2); }
		}

		/// <summary>
		/// Checks if string has Extended record
		/// </summary>
		public bool HasExtString
		{
			get { return false; }
			// ((Flags & FormattedUnicodeStringFlags.HasExtendedString) == FormattedUnicodeStringFlags.HasExtendedString); }
		}

		/// <summary>
		/// Checks if string has formatting
		/// </summary>
		public bool HasFormatting
		{
			get { return ((Flags & FormattedUnicodeStringFlags.HasFormatting) == FormattedUnicodeStringFlags.HasFormatting); }
		}

		/// <summary>
		/// Checks if string is unicode
		/// </summary>
		public bool IsMultiByte
		{
			get { return ((Flags & FormattedUnicodeStringFlags.MultiByte) == FormattedUnicodeStringFlags.MultiByte); }
		}

		/// <summary>
		/// Returns length of string in bytes
		/// </summary>
		private uint ByteCount
		{
			get { return (uint)(CharacterCount * ((IsMultiByte) ? 2 : 1)); }
		}

		/// <summary>
		/// Returns number of formats used for formatting (0 if string has no formatting)
		/// </summary>
		public ushort FormatCount
		{
			get { return (HasFormatting) ? BitConverter.ToUInt16(m_bytes, (int)m_offset + 3) : (ushort)0; }
		}

		/// <summary>
		/// Returns size of extended string in bytes, 0 if there is no one
		/// </summary>
		public uint ExtendedStringSize
		{
			get { return (HasExtString) ? (uint)BitConverter.ToUInt16(m_bytes, (int)m_offset + ((HasFormatting) ? 5 : 3)) : 0; }
		}

		/// <summary>
		/// Returns head (before string data) size in bytes
		/// </summary>
		public uint HeadSize
		{
			get { return (uint)((HasFormatting) ? 2 : 0) + (uint)((HasExtString) ? 4 : 0) + 3; }
		}

		/// <summary>
		/// Returns tail (after string data) size in bytes
		/// </summary>
		public uint TailSize
		{
			get { return (uint)((HasFormatting) ? 4 * FormatCount : 0) + ((HasExtString) ? ExtendedStringSize : 0); }
		}

		/// <summary>
		/// Returns size of whole record in bytes
		/// </summary>
		public uint Size
		{
			get
			{
				uint extraSize = (uint)((HasFormatting) ? (2 + FormatCount * 4) : 0) +
								 ((HasExtString) ? (4 + ExtendedStringSize) : 0) + 3;
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
				return IsMultiByte ?
					Encoding.Unicode.GetString(m_bytes, (int)(m_offset + HeadSize), (int)ByteCount) :
					 Encoding.Default.GetString(m_bytes, (int)(m_offset + HeadSize), (int)ByteCount);
			}
		}
	}
}