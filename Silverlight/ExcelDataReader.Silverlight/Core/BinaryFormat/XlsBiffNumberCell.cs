namespace ExcelDataReader.Silverlight.Core.BinaryFormat
{
	/// <summary>
	/// Represents a floating-point number 
	/// </summary>
	internal class XlsBiffNumberCell : XlsBiffBlankCell
	{
		internal XlsBiffNumberCell(byte[] bytes)
			: this(bytes, 0)
		{
		}

		internal XlsBiffNumberCell(byte[] bytes, uint offset)
			: base(bytes, offset)
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