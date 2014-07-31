namespace Excel.Core.BinaryFormat
{
	/// <summary>
	/// Represents a constant integer number in range 0..65535
	/// </summary>
	internal class XlsBiffIntegerCell : XlsBiffBlankCell
	{
		internal XlsBiffIntegerCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
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