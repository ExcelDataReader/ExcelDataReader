namespace ExcelDataReader.Silverlight.Core.BinaryFormat
{
	using System.IO;
	using Silverlight.Core.BinaryFormat;

	/// <summary>
	/// Represents an Excel file stream
	/// </summary>
	internal class XlsStream
	{
		protected XlsFat m_fat;
		protected Stream m_fileStream;
		protected XlsHeader m_hdr;
		protected uint m_startSector;

		public XlsStream(XlsHeader hdr, uint startSector)
		{
			m_fileStream = hdr.FileStream;
			m_fat = hdr.FAT;
			m_hdr = hdr;
			m_startSector = startSector;
		}

		/// <summary>
		/// Returns offset of first stream sector
		/// </summary>
		public uint BaseOffset
		{
			get { return (uint)((m_startSector + 1) * m_hdr.SectorSize); }
		}

		/// <summary>
		/// Returns number of first stream sector
		/// </summary>
		public uint BaseSector
		{
			get { return (m_startSector); }
		}

		/// <summary>
		/// Reads stream data from file
		/// </summary>
		/// <returns>Stream data</returns>
		public byte[] ReadStream()
		{

			uint sector = m_startSector, prevSector = 0;
			int sectorSize = m_hdr.SectorSize;

			byte[] buff = new byte[sectorSize];
			byte[] ret;

			using (MemoryStream ms = new MemoryStream(sectorSize * 8))
			{
				lock (m_fileStream)
				{
					do
					{
						if (prevSector == 0 || (sector - prevSector) != 1)
							m_fileStream.Seek((sector + 1) * sectorSize, SeekOrigin.Begin);
						prevSector = sector;
						m_fileStream.Read(buff, 0, sectorSize);
						ms.Write(buff, 0, sectorSize);
					} while ((sector = m_fat.GetNextSector(sector)) != (uint)FATMARKERS.FAT_EndOfChain);
				}

				ret = ms.ToArray();
			}

			return ret;
		}
	}
}
