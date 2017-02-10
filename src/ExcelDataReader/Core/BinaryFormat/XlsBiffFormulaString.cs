using System.Text;

using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents a string value of formula
	/// </summary>
	internal class XlsBiffFormulaString : XlsBiffRecord
	{
	    private readonly XlsFormattedUnicodeString m_unicodeString;

		internal XlsBiffFormulaString(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
		{
		    m_unicodeString = new XlsFormattedUnicodeString(bytes, offset + 4); 
		}

	    /// <summary>
		/// String text
		/// </summary>
		public string Value => m_unicodeString.Value;
	}
}