namespace ExcelDataReader.Silverlight.Core.BinaryFormat
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using Silverlight.Core.BinaryFormat;

	/// <summary>
	/// Represents Excel file FAT table
	/// </summary>
	internal class XlsFat
	{
		private readonly List<uint> m_fat;
		private readonly XlsHeader m_hdr;
		private readonly int m_sectors;
		private readonly int m_sectors_for_fat;

		/// <summary>
		/// Constructs FAT table from list of sectors
		/// </summary>
		/// <param name="hdr">XlsHeader</param>
		/// <param name="sectors">Sectors list</param>
		public XlsFat(XlsHeader hdr, List<uint> sectors)
		{
			m_hdr = hdr;
			m_sectors_for_fat = sectors.Count;
			uint sector = 0, prevSector = 0;
			int sectorSize = hdr.SectorSize;
			byte[] buff = new byte[sectorSize];
			Stream file = hdr.FileStream;
			using (MemoryStream ms = new MemoryStream(sectorSize * m_sectors_for_fat))
			{
				lock (file)
				{
					for (int i = 0; i < sectors.Count; i++)
					{
						sector = sectors[i];
						if (prevSector == 0 || (sector - prevSector) != 1)
							file.Seek((sector + 1) * sectorSize, SeekOrigin.Begin);
						prevSector = sector;
						file.Read(buff, 0, sectorSize);
						ms.Write(buff, 0, sectorSize);
					}
				}
				ms.Seek(0, SeekOrigin.Begin);
				BinaryReader rd = new BinaryReader(ms);
				m_sectors = (int)ms.Length / 4;
				m_fat = new List<uint>(m_sectors);
				for (int i = 0; i < m_sectors; i++)
					m_fat.Add(rd.ReadUInt32());
				rd.Close();
				ms.Close();
			}
		}

		/// <summary>
		/// Resurns number of sectors used by FAT itself
		/// </summary>
		public int SectorsForFat
		{
			get { return m_sectors_for_fat; }
		}

		/// <summary>
		/// Returns number of sectors described by FAT
		/// </summary>
		public int SectorsCount
		{
			get { return m_sectors; }
		}

		/// <summary>
		/// Returns underlying XlsHeader object
		/// </summary>
		public XlsHeader Header
		{
			get { return m_hdr; }
		}

		/// <summary>
		/// Returns next data sector using FAT
		/// </summary>
		/// <param name="sector">Current data sector</param>
		/// <returns>Next data sector</returns>
		public uint GetNextSector(uint sector)
		{
			if (m_fat.Count <= sector)
				throw new ArgumentException(Errors.ErrorFATBadSector);
			uint value = m_fat[(int)sector];
			if (value == (uint)FATMARKERS.FAT_FatSector || value == (uint)FATMARKERS.FAT_DifSector)
				throw new InvalidOperationException(Errors.ErrorFATRead);
			return value;
		}
	}
}
