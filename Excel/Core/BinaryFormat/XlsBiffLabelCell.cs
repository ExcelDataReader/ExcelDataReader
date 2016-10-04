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
        //private Encoding m_UseEncoding = Encoding.Unicode;
	    private IXlsString xlsString;

	    internal XlsBiffLabelCell(byte[] bytes, uint offset, uint stringOffset, ExcelBinaryReader reader)
	        : base(bytes, offset, reader)
	    {
            xlsString = XlsStringFactory.CreateXlsString(bytes, offset + stringOffset, reader);
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
                return xlsString.Value;
			}
		}
	}
}