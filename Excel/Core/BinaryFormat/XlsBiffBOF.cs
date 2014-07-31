namespace Excel.Core.BinaryFormat
{
	/// <summary>
	/// Represents BIFF BOF record
	/// </summary>
	internal class XlsBiffBOF : XlsBiffRecord
	{
		internal XlsBiffBOF(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
		{
		}

		/// <summary>
		/// Version
		/// </summary>
		public ushort Version
		{
			get { return ReadUInt16(0x0); }
		}

		/// <summary>
		/// Type of BIFF block
		/// </summary>
		public BIFFTYPE Type
		{
			get { return (BIFFTYPE) ReadUInt16(0x2); }
		}

		/// <summary>
		/// Creation ID
		/// </summary>
		/// <remarks>Not used before BIFF5</remarks>
		public ushort CreationID
		{
			get
			{
				if (RecordSize < 6) return 0;
				return ReadUInt16(0x4);
			}
		}

		/// <summary>
		/// Creation year
		/// </summary>
		/// <remarks>Not used before BIFF5</remarks>
		public ushort CreationYear
		{
			get
			{
				if (RecordSize < 8) return 0;
				return ReadUInt16(0x6);
			}
		}

		/// <summary>
		/// File history flag
		/// </summary>
		/// <remarks>Not used before BIFF8</remarks>
		public uint HistoryFlag
		{
			get
			{
				if (RecordSize < 12) return 0;
				return ReadUInt32(0x8);
			}
		}

		/// <summary>
		/// Minimum Excel version to open this file
		/// </summary>
		/// <remarks>Not used before BIFF8</remarks>
		public uint MinVersionToOpen
		{
			get
			{
				if (RecordSize < 16) return 0;
				return ReadUInt32(0xC);
			}
		}
	}
}