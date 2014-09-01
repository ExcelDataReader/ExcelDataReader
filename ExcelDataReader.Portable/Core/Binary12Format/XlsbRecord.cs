using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Excel.Core.Binary12Format
{
	/// <summary>
	/// Represents basic BIFF12 record
	/// Base class for all BIFF12 record types
	/// </summary>
	internal class XlsbRecord
	{
		#region Members and Properties

		protected byte[] _bytes;
		/// <summary>
		/// read offset
		/// </summary>
		protected int _offset;

		public BIFF12 ID
		{
			get { return GetID(_bytes, _offset); }
		}

		private static BIFF12 GetID(byte[] bytes, int offset)
		{
			BIFF12 val1 = (BIFF12)BitConverter.ToUInt16(bytes, offset - 4);

			BIFF12 val2 = BIFF12.UNKNOWN;

			uint recid = 0;

			if (offset < bytes.Length)
			{

				byte b1 = bytes[offset++];
				recid = (UInt32)(b1);

				if ((b1 & 0x80) > 0)
				{
					if (offset < bytes.Length)
					{
						byte b2 = bytes[offset++];
						recid = ((UInt32)(b2) << 8) | recid;

						if ((b2 & 0x80) == 0)
							return (BIFF12)b2;

						if (offset >= bytes.Length)
							return BIFF12.UNKNOWN;
						byte b3 = bytes[offset++];
						recid = ((UInt32)(b3) << 16) | recid;

						if ((b3 & 0x80) == 0)
							return (BIFF12)b3; ;

						if (offset >= bytes.Length)
							return BIFF12.UNKNOWN;
						byte b4 = bytes[offset++];
						recid = ((UInt32)(b4) << 24) | recid;

						return (BIFF12)b4;
					}
				}
			}

			return val2;
		}

		public uint GetLength()
		{
			uint reclen = 0;

			if (_offset >= _bytes.Length)
				return 0;

			byte b1 = _bytes[_offset++];
			reclen = (UInt32)(b1 & 0x7F);

			if ((b1 & 0x80) == 0)
				return reclen;

			if (_offset >= _bytes.Length)
				return 0;

			byte b2 = _bytes[_offset++];
			reclen = ((UInt32)(b2 & 0x7F) << 7) | reclen;

			if ((b2 & 0x80) == 0)
				return reclen;

			if (_offset >= _bytes.Length)
				return 0;

			byte b3 = _bytes[_offset++];
			reclen = ((UInt32)(b3 & 0x7F) << 14) | reclen;

			if ((b3 & 0x80) == 0)
				return reclen;

			if (_offset >= _bytes.Length)
				return 0;

			byte b4 = _bytes[_offset++];
			reclen = ((UInt32)(b4 & 0x7F) << 21) | reclen;

			return reclen;
		}

		/// <summary>
		/// Gets the record bytes.
		/// </summary>
		/// <value>The record bytes.</value>
		internal byte[] Bytes
		{
			get { return _bytes; }
		}

		/// <summary>
		/// Gets the read offset.
		/// </summary>
		/// <value>The read offset.</value>
		internal int Offset
		{
			get { return _offset - 4; }
		}

		public ushort RecordSize
		{
			get { return BitConverter.ToUInt16(_bytes, _offset - 2); }
		}

		public int Size
		{
			get { return 4 + RecordSize; }
		}

		#endregion

		protected XlsbRecord(byte[] bytes)
			: this(bytes, 0)
		{
		}

		protected XlsbRecord(byte[] bytes, uint offset)
		{
			if (bytes.Length - offset < 4)
				throw new ArgumentException(Errors.ErrorBIFFRecordSize);

			_bytes = bytes;
			_offset = (int)(4 + offset);

			//if (_bytes.Length < _offset + Size)//Size
			//    throw new ArgumentException(Errors.ErrorBIFFBufferSize);
		}


		public static XlsbRecord GetRecord(byte[] bytes, uint offset)
		{
			return new XlsbRecord(bytes, offset);
		}

		public int ReadUInt32(int offset)
		{
			int val1 = (int)BitConverter.ToUInt32(_bytes, _offset + offset);

			UInt32 val2 = ((UInt32)(_bytes[_offset + offset + 3]) << 24) +
				((UInt32)(_bytes[_offset + offset + 2]) << 16) +
				((UInt32)(_bytes[_offset + offset + 1]) << 8) +
				((UInt32)(_bytes[_offset + offset + 0]));

			System.Diagnostics.Debug.Assert(val1 == (int)val2);

			return (int)val2;
		}

		public ushort ReadUInt16(int offset)
		{
			ushort val1 = BitConverter.ToUInt16(_bytes, _offset + offset);

			UInt16 val = (UInt16)(_bytes[_offset + offset + 1] << 8);
			val += (UInt16)(_bytes[_offset + offset + 0]);

			System.Diagnostics.Debug.Assert(val1 == val);

			return val;
		}

		public byte ReadByte(int offset)
		{
			byte val1 = Buffer.GetByte(_bytes, _offset + offset);

			byte val2 = _bytes[_offset + offset];

			System.Diagnostics.Debug.Assert(val1 == val2);

			return val2;
		}

		public string ReadString(int offset, int len)
		{
			//StringBuilder sb = new StringBuilder((int)len);
			//for (UInt32 i = offset; i < offset + 2 * len; i += 2)
			//    sb.Append((Char)GetWord(buffer, i));
			//return sb.ToString();

			throw new NotImplementedException();
		}

		public double ReadFloat(int offset)
		{
			double d = 0;

			// When it's a simple precision float, Excel uses a special
			// encoding
			int rk = ReadUInt32(_offset + offset);
			if ((rk & 0x02) != 0)
			{
				// int
				d = (double)(rk >> 2);
			}
			else
			{
				using (MemoryStream mem = new MemoryStream())
				{
					BinaryWriter bw = new BinaryWriter(mem);

					bw.Write(0);
					bw.Write(rk & -4);

					mem.Seek(0, SeekOrigin.Begin);

					BinaryReader br = new BinaryReader(mem);
					d = br.ReadDouble();
					br.Close();
					bw.Close();
				}
			}
			if ((rk & 0x01) != 0)
			{
				// divide by 100
				d /= 100;
			}

			float val1 = BitConverter.ToSingle(_bytes, _offset + offset);

			System.Diagnostics.Debug.Assert(val1 == (float)d);

			return d;
		}

		public double ReadDouble(int offset)
		{
			double d = 0;

			using (MemoryStream mem = new MemoryStream())
			{
				BinaryWriter bw = new BinaryWriter(mem);

				for (UInt32 i = 0; i < 8; i++)
					bw.Write(_bytes[offset + _offset + i]);

				mem.Seek(0, SeekOrigin.Begin);

				BinaryReader br = new BinaryReader(mem);
				d = br.ReadDouble();
				br.Close();
				bw.Close();
			}

			double val1 = BitConverter.ToDouble(_bytes, _offset + offset);

			System.Diagnostics.Debug.Assert(val1 == d);

			return d;
		}

		//public static bool GetRecordID(byte[] buffer, ref UInt32 offset, ref UInt32 recid)
		//{
		//    recid = 0;

		//    if (offset >= buffer.Length)
		//        return false;
		//    byte b1 = buffer[offset++];
		//    recid = (UInt32)(b1);

		//    if ((b1 & 0x80) == 0)
		//        return true;

		//    if (offset >= buffer.Length)
		//        return false;
		//    byte b2 = buffer[offset++];
		//    recid = ((UInt32)(b2) << 8) | recid;

		//    if ((b2 & 0x80) == 0)
		//        return true;

		//    if (offset >= buffer.Length)
		//        return false;
		//    byte b3 = buffer[offset++];
		//    recid = ((UInt32)(b3) << 16) | recid;

		//    if ((b3 & 0x80) == 0)
		//        return true;

		//    if (offset >= buffer.Length)
		//        return false;
		//    byte b4 = buffer[offset++];
		//    recid = ((UInt32)(b4) << 24) | recid;

		//    return true;
		//}

		//public static bool GetRecordLen(byte[] buffer, ref UInt32 offset, ref UInt32 reclen)
		//{
		//    reclen = 0;

		//    if (offset >= buffer.Length)
		//        return false;
		//    byte b1 = buffer[offset++];
		//    reclen = (UInt32)(b1 & 0x7F);

		//    if ((b1 & 0x80) == 0)
		//        return true;

		//    if (offset >= buffer.Length)
		//        return false;
		//    byte b2 = buffer[offset++];
		//    reclen = ((UInt32)(b2 & 0x7F) << 7) | reclen;

		//    if ((b2 & 0x80) == 0)
		//        return true;

		//    if (offset >= buffer.Length)
		//        return false;
		//    byte b3 = buffer[offset++];
		//    reclen = ((UInt32)(b3 & 0x7F) << 14) | reclen;

		//    if ((b3 & 0x80) == 0)
		//        return true;

		//    if (offset >= buffer.Length)
		//        return false;
		//    byte b4 = buffer[offset++];
		//    reclen = ((UInt32)(b4 & 0x7F) << 21) | reclen;

		//    return true;
		//}
	}
}
