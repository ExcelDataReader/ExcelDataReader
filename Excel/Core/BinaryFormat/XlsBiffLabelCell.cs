using System.Text;
#if LEGACY
using Excel;
#endif
namespace ExcelDataReader.Portable.Core.BinaryFormat
{
	/// <summary>
	/// Represents a string (max 255 bytes)
	/// </summary>
	internal class XlsBiffLabelCell : XlsBiffBlankCell
	{
        private Encoding m_UseEncoding = Encoding.Unicode;
	    private XlsFormattedUnicodeString unicodeString;

	    internal XlsBiffLabelCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
	    {
	        unicodeString = new XlsFormattedUnicodeString(bytes, offset + 10);
	    }

	    /// <summary>
		/// Encoding used to deal with strings
		/// </summary>
		public Encoding UseEncoding
		{
			get { return unicodeString.UseEncoding; }
		}

		/// <summary>
		/// Length of string value
		/// </summary>
		public ushort Length
		{
            get { return unicodeString.CharacterCount; }
		}

		/// <summary>
		/// Returns value of this cell
		/// </summary>
		public string Value
		{
			get
			{
			    return unicodeString.Value;
			    //byte[] bts;

			    //if (reader.isV8())
			    //{
			    //    //issue 11636 - according to spec character data starts at byte 9 for biff8 (was using 8)
			    //    bts = base.ReadArray(0x9, Length * (Helpers.IsSingleByteEncoding(m_UseEncoding) ? 1 : 2));
			    //}
			    //else
			    //{ //biff 3-5
			    //    bts = base.ReadArray(0x2, Length * (Helpers.IsSingleByteEncoding(m_UseEncoding) ? 1 : 2));
			    //}


			    //return m_UseEncoding.GetString(bts, 0, bts.Length);
			}
		}
	}
}