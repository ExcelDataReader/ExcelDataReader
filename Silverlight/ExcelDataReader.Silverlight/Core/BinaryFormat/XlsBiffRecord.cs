namespace ExcelDataReader.Silverlight.Core.BinaryFormat
{
	using System;

	/// <summary>
	/// Represents basic BIFF record
	/// Base class for all BIFF record types
	/// </summary>
	internal class XlsBiffRecord
	{
		protected byte[] m_bytes;
		protected int m_readoffset;

		/// <summary>
		/// Basic entry constructor
		/// </summary>
		/// <param name="bytes">array representing this entry</param>
		protected XlsBiffRecord(byte[] bytes)
			: this(bytes, 0)
		{
		}

		protected XlsBiffRecord(byte[] bytes, uint offset)
		{
			if (bytes.Length - offset < 4)
				throw new ArgumentException(Errors.ErrorBIFFRecordSize);
			m_bytes = bytes;
			m_readoffset = (int)(4 + offset);
			if (bytes.Length < offset + Size)
				throw new ArgumentException(Errors.ErrorBIFFBufferSize);
		}

		internal byte[] Bytes
		{
			get { return m_bytes; }
		}

		internal int Offset
		{
			get { return m_readoffset - 4; }
		}

		/// <summary>
		/// Returns type ID of this entry
		/// </summary>
		public BIFFRECORDTYPE ID
		{
			get { return (BIFFRECORDTYPE)BitConverter.ToUInt16(m_bytes, m_readoffset - 4); }
		}

		/// <summary>
		/// Returns data size of this entry
		/// </summary>
		public ushort RecordSize
		{
			get { return BitConverter.ToUInt16(m_bytes, m_readoffset - 2); }
		}

		/// <summary>
		/// Returns whole size of structure
		/// </summary>
		public int Size
		{
			get { return 4 + RecordSize; }
		}

		/// <summary>
		/// Returns record at specified offset
		/// </summary>
		/// <param name="bytes">byte array</param>
		/// <param name="offset">position in array</param>
		/// <returns></returns>
		public static XlsBiffRecord GetRecord(byte[] bytes, uint offset)
		{
			uint ID = BitConverter.ToUInt16(bytes, (int)offset);
			switch ((BIFFRECORDTYPE)ID)
			{
				case BIFFRECORDTYPE.BOF_V2:
				case BIFFRECORDTYPE.BOF_V3:
				case BIFFRECORDTYPE.BOF_V4:
				case BIFFRECORDTYPE.BOF:
					return new XlsBiffBOF(bytes, offset);
				case BIFFRECORDTYPE.EOF:
					return new XlsBiffEOF(bytes, offset);
				case BIFFRECORDTYPE.INTERFACEHDR:
					return new XlsBiffInterfaceHdr(bytes, offset);

				case BIFFRECORDTYPE.SST:
					return new XlsBiffSST(bytes, offset);

				case BIFFRECORDTYPE.INDEX:
					return new XlsBiffIndex(bytes, offset);
				case BIFFRECORDTYPE.ROW:
					return new XlsBiffRow(bytes, offset);
				case BIFFRECORDTYPE.DBCELL:
					return new XlsBiffDbCell(bytes, offset);

				case BIFFRECORDTYPE.BLANK:
				case BIFFRECORDTYPE.BLANK_OLD:
					return new XlsBiffBlankCell(bytes, offset);
				case BIFFRECORDTYPE.MULBLANK:
					return new XlsBiffMulBlankCell(bytes, offset);
				case BIFFRECORDTYPE.LABEL:
				case BIFFRECORDTYPE.LABEL_OLD:
				case BIFFRECORDTYPE.RSTRING:
					return new XlsBiffLabelCell(bytes, offset);
				case BIFFRECORDTYPE.LABELSST:
					return new XlsBiffLabelSSTCell(bytes, offset);
				case BIFFRECORDTYPE.INTEGER:
				case BIFFRECORDTYPE.INTEGER_OLD:
					return new XlsBiffIntegerCell(bytes, offset);
				case BIFFRECORDTYPE.NUMBER:
				case BIFFRECORDTYPE.NUMBER_OLD:
					return new XlsBiffNumberCell(bytes, offset);
				case BIFFRECORDTYPE.RK:
					return new XlsBiffRKCell(bytes, offset);
				case BIFFRECORDTYPE.MULRK:
					return new XlsBiffMulRKCell(bytes, offset);
				case BIFFRECORDTYPE.FORMULA:
				case BIFFRECORDTYPE.FORMULA_OLD:
					return new XlsBiffFormulaCell(bytes, offset);
				case BIFFRECORDTYPE.STRING:
					return new XlsBiffFormulaString(bytes, offset);
				case BIFFRECORDTYPE.CONTINUE:
					return new XlsBiffContinue(bytes, offset);
				case BIFFRECORDTYPE.DIMENSIONS:
					return new XlsBiffDimensions(bytes, offset);
				case BIFFRECORDTYPE.BOUNDSHEET:
					return new XlsBiffBoundSheet(bytes, offset);
				case BIFFRECORDTYPE.WINDOW1:
					return new XlsBiffWindow1(bytes, offset);
				case BIFFRECORDTYPE.CODEPAGE:
					return new XlsBiffSimpleValueRecord(bytes, offset);
				case BIFFRECORDTYPE.FNGROUPCOUNT:
					return new XlsBiffSimpleValueRecord(bytes, offset);
				case BIFFRECORDTYPE.RECORD1904:
					return new XlsBiffSimpleValueRecord(bytes, offset);
				case BIFFRECORDTYPE.BOOKBOOL:
					return new XlsBiffSimpleValueRecord(bytes, offset);
				case BIFFRECORDTYPE.BACKUP:
					return new XlsBiffSimpleValueRecord(bytes, offset);
				case BIFFRECORDTYPE.HIDEOBJ:
					return new XlsBiffSimpleValueRecord(bytes, offset);
				case BIFFRECORDTYPE.USESELFS:
					return new XlsBiffSimpleValueRecord(bytes, offset);

				default:
					return new XlsBiffRecord(bytes, offset);
			}
		}

		public byte ReadByte(int offset)
		{
			return Buffer.GetByte(m_bytes, m_readoffset + offset);
		}

		public ushort ReadUInt16(int offset)
		{
			return BitConverter.ToUInt16(m_bytes, m_readoffset + offset);
		}

		public uint ReadUInt32(int offset)
		{
			return BitConverter.ToUInt32(m_bytes, m_readoffset + offset);
		}

		public ulong ReadUInt64(int offset)
		{
			return BitConverter.ToUInt64(m_bytes, m_readoffset + offset);
		}

		public short ReadInt16(int offset)
		{
			return BitConverter.ToInt16(m_bytes, m_readoffset + offset);
		}

		public int ReadInt32(int offset)
		{
			return BitConverter.ToInt32(m_bytes, m_readoffset + offset);
		}

		public long ReadInt64(int offset)
		{
			return BitConverter.ToInt64(m_bytes, m_readoffset + offset);
		}

		public byte[] ReadArray(int offset, int size)
		{
			byte[] tmp = new byte[size];
			Buffer.BlockCopy(m_bytes, m_readoffset + offset, tmp, 0, size);
			return tmp;
		}

		public float ReadFloat(int offset)
		{
			return BitConverter.ToSingle(m_bytes, m_readoffset + offset);
		}

		public double ReadDouble(int offset)
		{
			return BitConverter.ToDouble(m_bytes, m_readoffset + offset);
		}
	}
}
