
using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents Dimensions of worksheet
	/// </summary>
	internal class XlsBiffDimensions : XlsBiffRecord
	{
		private bool isV8 = true;

		internal XlsBiffDimensions(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
		{
		}

		/// <summary>
		/// Gets or sets if BIFF8 addressing is used
		/// </summary>
		public bool IsV8
		{
			get { return isV8; }
			set { isV8 = value; }
		}

		/// <summary>
		/// Index of first row
		/// </summary>
		public uint FirstRow
		{
			get { return (isV8) ? base.ReadUInt32(0x0) : base.ReadUInt16(0x0); }
		}

		/// <summary>
		/// Index of last row + 1
		/// </summary>
		public uint LastRow
		{
			get { return (isV8) ? base.ReadUInt32(0x4) : base.ReadUInt16(0x2); }
		}

		/// <summary>
		/// Index of first column
		/// </summary>
		public ushort FirstColumn
		{
			get { return (isV8) ? base.ReadUInt16(0x8) : base.ReadUInt16(0x4); }
		}

		/// <summary>
		/// Index of last column + 1
		/// </summary>
		public ushort LastColumn
		{
			get { return (isV8) ? (ushort)((base.ReadUInt16(0x9) >> 8) + 1) : base.ReadUInt16(0x6); }
			set { throw new System.NotImplementedException(); }
		}
	}
}