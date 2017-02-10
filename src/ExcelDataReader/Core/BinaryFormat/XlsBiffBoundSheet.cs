using System.Text;

using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents Sheet record in Workbook Globals
	/// </summary>
	internal class XlsBiffBoundSheet : XlsBiffRecord
	{
		#region SheetType enum

		public enum SheetType : byte
		{
			Worksheet = 0x0,
			MacroSheet = 0x1,
			Chart = 0x2,
			VBModule = 0x6
		}

		#endregion

		#region SheetVisibility enum

		public enum SheetVisibility : byte
		{
			Visible = 0x0,
			Hidden = 0x1,
			VeryHidden = 0x2
		}

		#endregion

	    internal XlsBiffBoundSheet(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
		{
		}
		/// <summary>
		/// Worksheet data start offset
		/// </summary>
		public uint StartOffset => ReadUInt32(0x0);

	    /// <summary>
		/// Type of worksheet
		/// </summary>
		public SheetType Type => (SheetType)ReadByte(0x5);

	    /// <summary>
		/// Visibility of worksheet
		/// </summary>
		public SheetVisibility VisibleState => (SheetVisibility)ReadByte(0x4);

	    /// <summary>
        /// Name of worksheet
        /// </summary>
        public string SheetName
        {
            get
            {
                ushort len = ReadByte(0x6);

                const int start = 0x8;
                if (!IsV8)
                    return reader.Encoding.GetString(m_bytes, m_readoffset + start, Helpers.IsSingleByteEncoding(reader.Encoding) ? len : len * 2);

                if (ReadByte(0x7) == 0)
                {
                    byte[] bytes = new byte[len * 2];
                    for (int i = 0; i < len; i++)
                    {
                        bytes[i * 2] = m_bytes[m_readoffset + start + i];
                    }

                    return Encoding.Unicode.GetString(bytes, 0, len * 2);
                }

                return Encoding.Unicode.GetString(m_bytes, m_readoffset + start, len * 2);
            }
        }

        /// <summary>
        /// Specifies if BIFF8 format should be used
        /// </summary>
        public bool IsV8 { get; set; } = true;
	}
}