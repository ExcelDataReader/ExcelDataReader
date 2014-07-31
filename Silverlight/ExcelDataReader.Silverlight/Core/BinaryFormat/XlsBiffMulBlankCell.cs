namespace ExcelDataReader.Silverlight.Core.BinaryFormat
{
	/// <summary>
	/// Represents multiple Blank cell
	/// </summary>
	internal class XlsBiffMulBlankCell : XlsBiffBlankCell
	{
		internal XlsBiffMulBlankCell(byte[] bytes)
			: this(bytes, 0)
		{
		}

		internal XlsBiffMulBlankCell(byte[] bytes, uint offset)
			: base(bytes, offset)
		{
		}

		/// <summary>
		/// Zero-based index of last described column
		/// </summary>
		public ushort LastColumnIndex
		{
			get { return base.ReadUInt16(RecordSize - 2); }
		}

		/// <summary>
		/// Returns format forspecified column, column must be between ColumnIndex and LastColumnIndex
		/// </summary>
		/// <param name="ColumnIdx">Index of column</param>
		/// <returns>Format</returns>
		public ushort GetXF(ushort ColumnIdx)
		{
			int ofs = 4 + 6*(ColumnIdx - ColumnIndex);
			if (ofs > RecordSize - 2)
				return 0;
			return base.ReadUInt16(ofs);
		}
	}
}