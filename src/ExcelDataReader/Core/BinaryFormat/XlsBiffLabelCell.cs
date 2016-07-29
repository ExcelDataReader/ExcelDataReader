using System.Text;

using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
	/// Represents a string (max 255 bytes)
	/// </summary>
	internal class XlsBiffLabelCell : XlsBiffBlankCell
	{
        //private Encoding m_UseEncoding = Encoding.Unicode;
        private Encoding m_UseEncoding;
	    private IXlsString xlsString;

	    internal XlsBiffLabelCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
	        : base(bytes, offset, reader)
	    {
	        m_UseEncoding = reader.DefaultEncoding;
            xlsString = XlsStringFactory.CreateXlsString(bytes, offset, reader);

	    }



	    /// <summary>
		/// Encoding used to deal with strings
		/// </summary>
		public Encoding UseEncoding
		{
            get { return reader.Encoding; }
		}


		/// <summary>
		/// Length of string value
		/// </summary>
		public ushort Length
		{
            get { return base.ReadUInt16(0x6); }
            //get { return xlsString.CharacterCount; }
		}

		/// <summary>
		/// Returns value of this cell
		/// </summary>
		public string Value
		{
			get
			{
                //return xlsString.Value;

			    byte[] bts;

                if (reader.isV8())
                {
                    //issue 11636 - according to spec character data starts at byte 9 for biff8 (was using 8)
                    bts = base.ReadArray(0x9, Length * (Helpers.IsSingleByteEncoding(xlsString.UseEncoding) ? 1 : 2));
                }
                else
                { //biff 3-5
                    bts = base.ReadArray(0x2, Length * (Helpers.IsSingleByteEncoding(xlsString.UseEncoding) ? 1 : 2));
                }


                return xlsString.UseEncoding.GetString(bts, 0, bts.Length);
			}
		}
	}
}