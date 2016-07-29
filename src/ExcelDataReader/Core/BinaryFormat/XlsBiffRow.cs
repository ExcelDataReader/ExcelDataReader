
using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents row record in table
	/// </summary>
	internal class XlsBiffRow : XlsBiffRecord
	{
		internal XlsBiffRow(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
		{
		}

		/// <summary>
		/// Zero-based index of row described
		/// </summary>
		public ushort RowIndex
		{
			get { return base.ReadUInt16(0x0); }
		}

		/// <summary>
		/// Index of first defined column
		/// </summary>
		public ushort FirstDefinedColumn
		{
			get { return base.ReadUInt16(0x2); }
		}

		/// <summary>
		/// Index of last defined column
		/// </summary>
		public ushort LastDefinedColumn
		{
			get { return base.ReadUInt16(0x4); }
		}

		/// <summary>
		/// Returns row height
		/// </summary>
		public uint RowHeight
		{
			get { return base.ReadUInt16(0x6); }
		}

		/// <summary>
		/// Returns row flags
		/// </summary>
		public ushort Flags
		{
			get { return base.ReadUInt16(0xC); }
		}

		/// <summary>
		/// Returns default format for this row
		/// </summary>
		public ushort XFormat
		{
			get { return base.ReadUInt16(0xE); }
		}
	}
}