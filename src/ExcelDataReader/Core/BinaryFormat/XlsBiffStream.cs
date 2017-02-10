using System;
using System.IO;
#if NET45 || NET20
using System.Security.Cryptography;
#endif

using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents a BIFF stream
	/// </summary>
	internal class XlsBiffStream
	{
		private readonly ExcelBinaryReader m_reader;
		private readonly byte[] m_bytes;

	    public XlsBiffStream(XlsHeader hdr, uint streamStart, bool isMini, XlsRootDirectory rootDir, ExcelBinaryReader reader)
		{
			m_reader = reader;
		    var xlsStream = new XlsStream(hdr, streamStart, isMini, rootDir);
            m_bytes = xlsStream.ReadStream();

#if NET45 || NET20
            XlsBiffRecord rec = XlsBiffRecord.GetRecord(m_bytes, 0, m_reader);
		    XlsBiffRecord rec2 = XlsBiffRecord.GetRecord(m_bytes, (uint)rec.Size, reader);
            XlsBiffFilePass filePass = rec2 as XlsBiffFilePass;
		    if (filePass == null)
		    {
                XlsBiffRecord rec3 = XlsBiffRecord.GetRecord(m_bytes, (uint)(rec.Size + rec2.Size), reader);
                filePass = rec3 as XlsBiffFilePass;
		    }

		    if (filePass != null)
		    {
		        RC4Key key = new RC4Key("VelvetSweatshop", filePass.Salt);

		        int blockNumber = 0;
		        RC4 rc4 = key.Create(blockNumber);

		        int position = 0;
		        while (position < m_bytes.Length - 4)
		        {
                    uint id = BitConverter.ToUInt16(m_bytes, position);
		            int length = BitConverter.ToUInt16(m_bytes, position + 2) + 4;

		            int startDecrypt = 4;
		            switch ((BIFFRECORDTYPE)id)
		            {
                        case BIFFRECORDTYPE.BOF:
                        case BIFFRECORDTYPE.FILEPASS:
                        case BIFFRECORDTYPE.INTERFACEHDR:
		                    startDecrypt = length;
                            break;
                        case BIFFRECORDTYPE.BOUNDSHEET:
		                    startDecrypt += 4; // For some reason the sheet offset is not encrypted
		                    break;
		            }

		            for (int i = 0; i < length; i++)
		            {
                        int currentBlock = position / 1024;
                        if (blockNumber != currentBlock)
                        {
                            blockNumber = currentBlock;
                            rc4 = key.Create(blockNumber);
                        }

                        byte mask = rc4.Output();
		                if (i >= startDecrypt)
		                {
		                    m_bytes[position] = (byte)(m_bytes[position] ^ mask);
		                }

		                position++;
		            }
		        }
		    }
#endif

            Size = m_bytes.Length;
			Position = 0;
		}

		/// <summary>
		/// Returns size of BIFF stream in bytes
		/// </summary>
		public int Size { get; }

	    /// <summary>
		/// Returns current position in BIFF stream
		/// </summary>
		public int Position { get; private set; }

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
                        Position = offset;
                        break;
                    case SeekOrigin.Current:
                        Position += offset;
                        break;
                    case SeekOrigin.End:
                        Position = Size - offset;
                        break;
                }
                if (Position < 0)
                    throw new ArgumentOutOfRangeException(string.Format("{0} On offset={1}", Errors.ErrorBIFFIlegalBefore, offset));
                if (Position > Size)
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
                // Minimum record size is 4
                if ((uint)Position + 4 >= m_bytes.Length)
                    return null;

                XlsBiffRecord rec = XlsBiffRecord.GetRecord(m_bytes, (uint)Position, m_reader);
                Position += rec.Size;
                if (Position > Size)
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
            if ((uint)offset >= m_bytes.Length)
                return null;

			XlsBiffRecord rec = XlsBiffRecord.GetRecord(m_bytes, (uint)offset, m_reader);

			//choose ReadOption.Loose to skip this check (e.g. sql reporting services)
			if (m_reader.ReadOption == ReadOption.Strict)
			{
				if (Position + rec.Size > Size)
					return null;
			}
			
			return rec;
		}

#if NET45 || NET20
        private sealed class RC4Key
	    {
	        private readonly byte[] m_key;

	        public RC4Key(string password, byte[] salt)
	        {
                int length = Math.Min(password.Length, 16);
                byte[] passwordData = new byte[length * 2];
                for (int i = 0; i < length; i++)
                {
                    char ch = password[i];
                    passwordData[i * 2 + 0] = (byte)((ch << 0) & 0xFF);
                    passwordData[i * 2 + 1] = (byte)((ch << 8) & 0xFF);
                }

	            using (MD5 md5 = new MD5CryptoServiceProvider())
                {
                    byte[] passwordHash = md5.ComputeHash(passwordData);

                    md5.Initialize();

                    const int truncateCount = 5;
                    byte[] intermediateData = new byte[truncateCount * 16 + salt.Length * 16];

                    int offset = 0;
                    for (int i = 0; i < 16; i++)
                    {
                        Array.Copy(passwordHash, 0, intermediateData, offset, truncateCount);
                        offset += truncateCount;
                        Array.Copy(salt, 0, intermediateData, offset, salt.Length);
                        offset += salt.Length;
                    }

                    const int keyLength = 5;

                    byte[] finalHash = md5.ComputeHash(intermediateData);
                    byte[] result = new byte[keyLength];
                    Array.Copy(finalHash, 0, result, 0, keyLength);
                    md5.Clear();

                    m_key = result;
                }
	        }


            public RC4 Create(int blockNumber)
            {
                byte[] data = new byte[4 + m_key.Length];
                data[data.Length - 1] = (byte)((blockNumber >> 24) & 0xFF);
                data[data.Length - 2] = (byte)((blockNumber >> 16) & 0xFF);
                data[data.Length - 3] = (byte)((blockNumber >> 8) & 0xFF);
                data[data.Length - 4] = (byte)((blockNumber >> 0) & 0xFF);

                Array.Copy(m_key, 0, data, 0, m_key.Length);

                using (MD5 md5 = new MD5CryptoServiceProvider())
                {
                    byte[] blockKey = md5.ComputeHash(data);

                    return new RC4(blockKey);
                }
            }
        }

	    private sealed class RC4
	    {
            private readonly byte[] m_s = new byte[256];

	        private int m_i;

	        private int m_j;

            public RC4(byte[] key)
	        {
	            for (int i = 0; i < m_s.Length; i++)
	            {
                    m_s[i] = (byte)i;
	            }

                for (int i = 0, j = 0; i < 256; i++)
                {
                    j = (j + key[i % key.Length] + m_s[i]) & 255;

                    Swap(m_s, i, j);
                }
            }

            public byte Output()
            {
                m_i = (m_i + 1) & 255;
                m_j = (m_j + m_s[m_i]) & 255;

                Swap(m_s, m_i, m_j);

                return m_s[(m_s[m_i] + m_s[m_j]) & 255];
            }

            private static void Swap(byte[] s, int i, int j)
            {
                byte c = s[i];

                s[i] = s[j];
                s[j] = c;
            }
        }
#endif
    }
}
