using System.Collections.Generic;

using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents a worksheet index
	/// </summary>
	internal class XlsBiffIndex : XlsBiffRecord
	{
	    internal XlsBiffIndex(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
		{
		    IsV8 = reader.IsV8();
		}

		/// <summary>
		/// Gets or sets if BIFF8 addressing is used
		/// </summary>
		public bool IsV8 { get; }

	    /// <summary>
		/// Returns zero-based index of first existing row
		/// </summary>
		public uint FirstExistingRow => IsV8 ? ReadUInt32(0x4) : ReadUInt16(0x4);

	    /// <summary>
		/// Returns zero-based index of last existing row
		/// </summary>
		public uint LastExistingRow => IsV8 ? ReadUInt32(0x8) : ReadUInt16(0x6);

	    /// <summary>
		/// Returns addresses of DbCell records
		/// </summary>
		public uint[] DbCellAddresses
		{
			get
			{
				int size = RecordSize;
				int firstIdx = (IsV8) ? 16 : 12;
				if (size <= firstIdx)
					return new uint[0];
				List<uint> cells = new List<uint>((size - firstIdx)/4);
				for (int i = firstIdx; i < size; i += 4)
					cells.Add(ReadUInt32(i));
				return cells.ToArray();
			}
		}
	}
}