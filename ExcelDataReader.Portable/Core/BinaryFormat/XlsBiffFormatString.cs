using System.Text;

namespace ExcelDataReader.Portable.Core.BinaryFormat
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
	    }


	    /// <summary>
        /// Encoding used to deal with strings
        /// </summary>
        public Encoding UseEncoding
        {
            get { return unicodeString.UseEncoding; }
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
                //if (m_value == null)
                //{
                //    switch (ID)
                //    {
                //        case BIFFRECORDTYPE.FORMAT_V23:
                //            m_value = m_UseEncoding.GetString(m_bytes, m_readoffset + 1, Length);
                //            break;
                //        case BIFFRECORDTYPE.FORMAT:
                //            var offset = m_readoffset + 5;
                //            var flags = ReadByte(3);
                //            m_UseEncoding = Encoding.Unicode; //we are assuming BIFF 8 so always unicode
                //            if ((flags & 0x04) == 0x01) // asian phonetic block size
                //                offset += 4;
                //            if ((flags & 0x08) == 0x01) // number of rtf blocks
                //                offset += 2;
                //            //m_value = m_UseEncoding.IsSingleByte ? m_UseEncoding.GetString(m_bytes, offset, Length) : m_UseEncoding.GetString(m_bytes, offset, Length*2);
                //            //note: BIFF8 is unicode always so don't need single byte check
                //            m_value = m_UseEncoding.GetString(m_bytes, offset, Length*2);

                //            break;


                //    }


                //}
                //return m_value;
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