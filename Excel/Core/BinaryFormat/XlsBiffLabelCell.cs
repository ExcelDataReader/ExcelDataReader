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
        private Encoding m_UseEncoding;
	    private IXlsString xlsString;

	    internal XlsBiffLabelCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
	        : base(bytes, offset, reader)
	    {
	        m_UseEncoding = reader.DefaultEncoding;
            // Label record consists of: record type (2 bytes)|record size (2 bytes)|Cell structure (6 bytes)|XLUnicodeString structure (variable)
            xlsString = XlsStringFactory.CreateXlsString(bytes, offset + 10, reader);

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
            get { return xlsString.CharacterCount; }
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