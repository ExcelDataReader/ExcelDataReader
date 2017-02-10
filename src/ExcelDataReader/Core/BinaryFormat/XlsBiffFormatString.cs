using System;
using System.Text;

using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents a string value of format
	/// </summary>
	internal class XlsBiffFormatString : XlsBiffRecord
	{
	    private readonly IXlsString m_string;

	    internal XlsBiffFormatString(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
	    {
            if (reader.IsV8())
	            m_string = new XlsFormattedUnicodeString(bytes, offset + 6);
	        else
	            m_string = new XlsByteString(bytes, offset + 4, reader.Encoding);
	    }

		/// <summary>
		/// String text
		/// </summary>
        public string Value => m_string.Value;

	    public ushort Index
        {
            get
            {
                switch (ID)
                {
                    case BIFFRECORDTYPE.FORMAT_V23:
                        throw new NotSupportedException("Index is not available for BIFF2 and BIFF3 FORMAT records.");
                    default:
                        return ReadUInt16(0);
                }
            }
        }
	}
}