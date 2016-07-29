using System.Text;

using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents a string value of format
	/// </summary>
	internal class XlsBiffFormatString : XlsBiffRecord
	{

        private Encoding m_UseEncoding =  Encoding.Unicode;
		private string m_value = null;
	    private XlsFormattedUnicodeString unicodeString;

	    internal XlsBiffFormatString(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
	    {
            unicodeString = new XlsFormattedUnicodeString(bytes, offset + 6);
            //unicodeString = new XlsFormattedUnicodeString(bytes, offset + 6, reader.Encoding);
	    }


	    /// <summary>
        /// Encoding used to deal with strings
        /// </summary>
        public Encoding UseEncoding
        {
            get { return reader.Encoding; }
            //set { m_UseEncoding = value; }
        }

		/// <summary>
		/// Length of the string
		/// </summary>
		public ushort Length
		{
			get
			{
			     switch (ID)
			     {
			         case BIFFRECORDTYPE.FORMAT_V23:
			             return base.ReadByte(0x0);
			         default:
			             return base.ReadUInt16(2);
			     }
			}
		}

		/// <summary>
		/// String text
		/// </summary>
        public string Value
        {
            get
            {
                return unicodeString.Value;

            }
        }

        public ushort Index
        {
            get
            {
                switch (ID)
                {
                    case BIFFRECORDTYPE.FORMAT_V23:
                        return 0;
                    default:
                        return ReadUInt16(0);

                }
            }
        }
	}
}