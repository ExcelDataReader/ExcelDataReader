using System;
using System.IO;

using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents a BIFF stream
	/// </summary>
	internal class XlsBiffStream : XlsStream
	{
		private readonly ExcelBinaryReader reader;
		private readonly byte[] bytes;
		private readonly int m_size;
		private int m_offset;

		public XlsBiffStream(XlsHeader hdr, uint streamStart, bool isMini, XlsRootDirectory rootDir, ExcelBinaryReader reader)
			: base(hdr, streamStart, isMini, rootDir)
		{
			this.reader = reader;
			bytes = base.ReadStream();
			m_size = bytes.Length;
			m_offset = 0;

		}

		/// <summary>
		/// Returns size of BIFF stream in bytes
		/// </summary>
		public int Size
		{
			get { return m_size; }
		}

		/// <summary>
		/// Returns current position in BIFF stream
		/// </summary>
		public int Position
		{
			get { return m_offset; }
		}

		//TODO:Remove ReadStream
		/// <summary>
		/// Always returns null, use biff-specific methods
		/// </summary>
		/// <returns></returns>
		[Obsolete("Use BIFF-specific methods for this stream")]
		public new byte[] ReadStream()
		{
			return bytes;
		}

		/// <summary>
		/// Sets stream pointer to the specified offset
		/// </summary>
		/// <param name="offset">Offset value</param>
		/// <param name="origin">Offset origin</param>
		public void Seek(int offset, SeekOrigin origin)
		{
            //add lock(this) as this is equivalent to [MethodImpl(MethodImplOptions.Synchronized)] on the method
            lock (this)
            {
                switch (origin)
                {
                    case SeekOrigin.Begin:
                        m_offset = offset;
                        break;
                    case SeekOrigin.Current:
                        m_offset += offset;
                        break;
                    case SeekOrigin.End:
                        m_offset = m_size - offset;
                        break;
                }
                if (m_offset < 0)
                    throw new ArgumentOutOfRangeException(string.Format("{0} On offset={1}", Errors.ErrorBIFFIlegalBefore, offset));
                if (m_offset > m_size)
                    throw new ArgumentOutOfRangeException(string.Format("{0} On offset={1}", Errors.ErrorBIFFIlegalAfter, offset));
            }

		}

		/// <summary>
		/// Reads record under cursor and advances cursor position to next record
		/// </summary>
		/// <returns></returns>
		public XlsBiffRecord Read()
		{
            //add lock(this) as this is equivalent to [MethodImpl(MethodImplOptions.Synchronized)] on the method
            lock (this)
            {
                if ((uint)m_offset >= bytes.Length)
                    return null;

                XlsBiffRecord rec = XlsBiffRecord.GetRecord(bytes, (uint)m_offset, reader);
                m_offset += rec.Size;
                if (m_offset > m_size)
                    return null;
                return rec;
            }

		}

		/// <summary>
		/// Reads record at specified offset, does not change cursor position
		/// </summary>
		/// <param name="offset"></param>
		/// <returns></returns>
		public XlsBiffRecord ReadAt(int offset)
		{
            if ((uint)offset >= bytes.Length)
                return null;

			XlsBiffRecord rec = XlsBiffRecord.GetRecord(bytes, (uint)offset, reader);

			//choose ReadOption.Loose to skip this check (e.g. sql reporting services)
			if (reader.ReadOption == ReadOption.Strict)
			{
				if (m_offset + rec.Size > m_size)
					return null;
			}
			
			return rec;
		}
	}
}
