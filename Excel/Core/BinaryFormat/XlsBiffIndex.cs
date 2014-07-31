using System.Collections.Generic;

namespace Excel.Core.BinaryFormat
{
	/// <summary>
	/// Represents a worksheet index
	/// </summary>
	internal class XlsBiffIndex : XlsBiffRecord
	{
		private bool isV8 = true;

		internal XlsBiffIndex(byte[] bytes, uint offset, ExcelBinaryReader reader)
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
		/// Returns zero-based index of first existing row
		/// </summary>
		public uint FirstExistingRow
		{
			get { return (isV8) ? base.ReadUInt32(0x4) : base.ReadUInt16(0x4); }
		}

		/// <summary>
		/// Returns zero-based index of last existing row
		/// </summary>
		public uint LastExistingRow
		{
			get { return (isV8) ? base.ReadUInt32(0x8) : base.ReadUInt16(0x6); }
		}

		/// <summary>
		/// Returns addresses of DbCell records
		/// </summary>
		public uint[] DbCellAddresses
		{
			get
			{
				int size = RecordSize;
				int firstIdx = (isV8) ? 16 : 12;
				if (size <= firstIdx)
					return new uint[0];
				List<uint> cells = new List<uint>((size - firstIdx)/4);
				for (int i = firstIdx; i < size; i += 4)
					cells.Add(base.ReadUInt32(i));
				return cells.ToArray();
			}
		}
	}
}