using System.Text;

using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
	/// Represents a string (max 255 bytes)
	/// </summary>
	internal class XlsBiffLabelCell : XlsBiffBlankCell
	{
	    private readonly IXlsString m_xlsString;

	    internal XlsBiffLabelCell(byte[] bytes, uint offset, uint stringOffset, ExcelBinaryReader reader)
	        : base(bytes, offset, reader)
	    {
            m_xlsString = XlsStringFactory.CreateXlsString(bytes, offset + stringOffset, reader);
	    }

		/// <summary>
		/// Length of string value
		/// </summary>
		public ushort Length => m_xlsString.CharacterCount;

	    /// <summary>
		/// Returns value of this cell
		/// </summary>
		public string Value => m_xlsString.Value;
	}
}