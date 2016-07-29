
using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents a floating-point number 
	/// </summary>
	internal class XlsBiffNumberCell : XlsBiffBlankCell
	{
		internal XlsBiffNumberCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
		{
		}

		/// <summary>
		/// Returns value of this cell
		/// </summary>
		public double Value
		{
			get { return base.ReadDouble(0x6); }
		}
	}
}