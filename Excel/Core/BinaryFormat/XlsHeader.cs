using System;
using System.Collections.Generic;
using System.IO;
using Excel.Exceptions;

namespace Excel.Core.BinaryFormat
{
	/// <summary>
	/// Represents Excel file header
	/// </summary>
	internal class XlsHeader
	{
		private readonly byte[] m_bytes;
		private readonly Stream m_file;
		private XlsFat m_fat;
		private XlsFat m_minifat;

		private XlsHeader(Stream file)
		{
			m_bytes = new byte[512];
			m_file = file;
		}

		/// <summary>
		/// Returns file signature
		/// </summary>
		public ulong Signature
		{
			get { return BitConverter.ToUInt64(m_bytes, 0x0); }
		}

		/// <summary>
		/// Checks if file signature is valid
		/// </summary>
		public bool IsSignatureValid
		{
			get { return (Signature == 0xE11AB1A1E011CFD0); }
		}

		/// <summary>
		/// Typically filled with zeroes
		/// </summary>
		public Guid ClassId
		{
			get
			{
				byte[] tmp = new byte[16];
				Buffer.BlockCopy(m_bytes, 0x8, tmp, 0, 16);
				return new Guid(tmp);
			}
		}

		/// <summary>
		/// Must be 0x003E
		/// </summary>
		public ushort Version
		{
			get { return BitConverter.ToUInt16(m_bytes, 0x18); }
		}

		/// <summary>
		/// Must be 0x0003
		/// </summary>
		public ushort DllVersion
		{
			get { return BitConverter.ToUInt16(m_bytes, 0x1A); }
		}

		/// <summary>
		/// Must be 0xFFFE
		/// </summary>
		public ushort ByteOrder
		{
			get { return BitConverter.ToUInt16(m_bytes, 0x1C); }
		}

		/// <summary>
		/// Typically 512
		/// </summary>
		public int SectorSize
		{
			get { return (1 << BitConverter.ToUInt16(m_bytes, 0x1E)); }
		}

		/// <summary>
		/// Typically 64
		/// </summary>
		public int MiniSectorSize
		{
			get { return (1 << BitConverter.ToUInt16(m_bytes, 0x20)); }
		}

		/// <summary>
		/// Number of FAT sectors
		/// </summary>
		public int FatSectorCount
		{
			get { return BitConverter.ToInt32(m_bytes, 0x2C); }
		}

		/// <summary>
		/// Number of first Root Directory Entry (Property Set Storage, FAT Directory) sector
		/// </summary>
		public uint RootDirectoryEntryStart
		{
			get { return BitConverter.ToUInt32(m_bytes, 0x30); }
		}

		/// <summary>
		/// Transaction signature, 0 for Excel
		/// </summary>
		public uint TransactionSignature
		{
			get { return BitConverter.ToUInt32(m_bytes, 0x34); }
		}

		/// <summary>
		/// Maximum size for small stream, typically 4096 bytes
		/// </summary>
		public uint MiniStreamCutoff
		{
			get { return BitConverter.ToUInt32(m_bytes, 0x38); }
		}

		/// <summary>
		/// First sector of Mini FAT, FAT_EndOfChain if there's no one
		/// </summary>
		public uint MiniFatFirstSector
		{
			get { return BitConverter.ToUInt32(m_bytes, 0x3C); }
		}

		/// <summary>
		/// Number of sectors in Mini FAT, 0 if there's no one
		/// </summary>
		public int MiniFatSectorCount
		{
			get { return BitConverter.ToInt32(m_bytes, 0x40); }
		}

		/// <summary>
		/// First sector of DIF, FAT_EndOfChain if there's no one
		/// </summary>
		public uint DifFirstSector
		{
			get { return BitConverter.ToUInt32(m_bytes, 0x44); }
		}

		/// <summary>
		/// Number of sectors in DIF, 0 if there's no one
		/// </summary>
		public int DifSectorCount
		{
			get { return BitConverter.ToInt32(m_bytes, 0x48); }
		}

		public Stream FileStream
		{
			get { return m_file; }
		}

		
		/// <summary>
		/// Returns mini FAT table
		/// </summary>
		public XlsFat GetMiniFAT(XlsRootDirectory rootDir)
		{
			if (m_minifat != null)
				return m_minifat;

			//if no minifat then return null
			if (MiniFatSectorCount == 0 || MiniSectorSize == 0xFFFFFFFE)
				return null;

			uint value;
			int miniSectorSize = MiniSectorSize;
			List<uint> sectors = new List<uint>(MiniFatSectorCount);

			//find the sector where the minifat starts
			var miniFatStartSector = BitConverter.ToUInt32(m_bytes, 0x3c);
			sectors.Add(miniFatStartSector);
			//lock (m_file)
			//{
			//	//work out the file location of minifat then read each sector
			//	var miniFatStartOffset = (miniFatStartSector + 1)*SectorSize;
			//	var miniFatSize = MiniFatSectorCount * SectorSize;
			//	m_file.Seek(miniFatStartOffset, SeekOrigin.Begin);

			//	byte[] sectorBuff = new byte[SectorSize];

			//	for (var i = 0; i < MiniFatSectorCount; i += SectorSize)
			//	{
			//		m_file.Read(sectorBuff, 0, 4);
			//		var secId = BitConverter.ToUInt32(sectorBuff, 0);
			//		sectors.Add(secId);
			//	}
			//}
				
			m_minifat = new XlsFat(this, sectors, this.MiniSectorSize, true, rootDir);
			return m_minifat;

		}

		/// <summary>
		/// Returns full FAT table, including DIF sectors
		/// </summary>
		public XlsFat FAT
		{
			get
			{
				if (m_fat != null)
					return m_fat;

				uint value;
				int sectorSize = SectorSize;
				List<uint> sectors = new List<uint>(FatSectorCount);
				for (int i = 0x4C; i < sectorSize; i += 4)
				{
					value = BitConverter.ToUInt32(m_bytes, i);
					if (value == (uint)FATMARKERS.FAT_FreeSpace)
						goto XlsHeader_Fat_Ready;
					sectors.Add(value);
				}
				int difCount;
				if ((difCount = DifSectorCount) == 0)
					goto XlsHeader_Fat_Ready;
				lock (m_file)
				{
					uint difSector = DifFirstSector;
					byte[] buff = new byte[sectorSize];
					uint prevSector = 0;
					while (difCount > 0)
					{
						sectors.Capacity += 128;
						if (prevSector == 0 || (difSector - prevSector) != 1)
							m_file.Seek((difSector + 1) * sectorSize, SeekOrigin.Begin);
						prevSector = difSector;
						m_file.Read(buff, 0, sectorSize);
						for (int i = 0; i < 508; i += 4)
						{
							value = BitConverter.ToUInt32(buff, i);
							if (value == (uint)FATMARKERS.FAT_FreeSpace)
								goto XlsHeader_Fat_Ready;
							sectors.Add(value);
						}
						value = BitConverter.ToUInt32(buff, 508);
						if (value == (uint)FATMARKERS.FAT_FreeSpace)
							break;
						if ((difCount--) > 1)
							difSector = value;
						else
							sectors.Add(value);
					}
				}
			XlsHeader_Fat_Ready:
				m_fat = new XlsFat(this, sectors, this.SectorSize, false, null);
				return m_fat;
			}
		}

		/// <summary>
		/// Reads Excel header from Stream
		/// </summary>
		/// <param name="file">Stream with Excel file</param>
		/// <returns>XlsHeader representing specified file</returns>
		public static XlsHeader ReadHeader(Stream file)
		{
			XlsHeader hdr = new XlsHeader(file);
			lock (file)
			{
				file.Seek(0, SeekOrigin.Begin);
				file.Read(hdr.m_bytes, 0, 512);
			}
			if (!hdr.IsSignatureValid)
				throw new HeaderException(Errors.ErrorHeaderSignature);
			if (hdr.ByteOrder != 0xFFFE)
				throw new FormatException(Errors.ErrorHeaderOrder);
			return hdr;
		}
	}
}
