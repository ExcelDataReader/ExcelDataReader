namespace ExcelDataReader.Silverlight.Core.BinaryFormat
{
	/// <summary>
	/// Represents a constant integer number in range 0..65535
	/// </summary>
	internal class XlsBiffIntegerCell : XlsBiffBlankCell
	{
		internal XlsBiffIntegerCell(byte[] bytes)
			: this(bytes, 0)
		{
		}

		internal XlsBiffIntegerCell(byte[] bytes, uint offset)
			: base(bytes, offset)
		{
		}

		/// <summary>
		/// Returns value of this cell
		/// </summary>
		public uint Value
		{
			get { return base.ReadUInt16(0x6); }
		}
	}
}