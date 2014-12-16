using System;
using System.IO;

namespace ExcelDataReader.Portable.Core.BinaryFormat
{
	/// <summary>
	/// Represents an Excel file stream
	/// </summary>
	internal class XlsStream
	{
		protected XlsFat m_fat;
		protected XlsFat m_minifat;
		protected Stream m_fileStream;
		protected XlsHeader m_hdr;
		protected uint m_startSector;
		protected bool m_isMini;
		protected XlsRootDirectory m_rootDir;

		public XlsStream(XlsHeader hdr, uint startSector, bool isMini, XlsRootDirectory rootDir)
		{
			m_fileStream = hdr.FileStream;
			m_fat = hdr.FAT;
			m_hdr = hdr;
			m_startSector = startSector;
			m_isMini = isMini;
			m_rootDir = rootDir;

			CalculateMiniFat(rootDir);

		}

		public void CalculateMiniFat(XlsRootDirectory rootDir)
		{
			m_minifat = m_hdr.GetMiniFAT(rootDir);
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
			int sectorSize = m_isMini ? m_hdr.MiniSectorSize : m_hdr.SectorSize;
			var fat = m_isMini ? m_minifat : m_fat;
			long offset = 0;
			if (m_isMini && m_rootDir != null)
			{
				offset = (m_rootDir.RootEntry.StreamFirstSector + 1)*m_hdr.SectorSize;
			}

			byte[] buff = new byte[sectorSize];
			byte[] ret;

			using (MemoryStream ms = new MemoryStream(sectorSize * 8))
			{
				lock (m_fileStream)
				{
					do
					{
						if (prevSector == 0 || (sector - prevSector) != 1)
						{
							var adjustedSector = m_isMini ? sector : sector + 1; //standard sector is + 1 because header is first
							m_fileStream.Seek(adjustedSector * sectorSize + offset, SeekOrigin.Begin);
						}

                        if (prevSector != 0 && prevSector == sector)
                            throw new InvalidOperationException("The excel file may be corrupt. We appear to be stuck");

						prevSector = sector;
						m_fileStream.Read(buff, 0, sectorSize);
						ms.Write(buff, 0, sectorSize);

					    sector = fat.GetNextSector(sector);

                        if (sector == 0)
                            throw new InvalidOperationException("Next sector cannot be 0. Possibly corrupt excel file");
					} while (sector != (uint)FATMARKERS.FAT_EndOfChain);
				}

				ret = ms.ToArray();
			}

			return ret;
		}
	}
}
