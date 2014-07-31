using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Excel.Core.Binary12Format
{
	internal class XlsbWorkbook
	{
		private XlsbWorkbook() { }

		public XlsbWorkbook(Stream workbookStream, Stream sharedStringsStream, Stream stylesStream)
		{
			if (null == workbookStream) throw new ArgumentNullException();

			ReadWorkbook(workbookStream);

			ReadSharedStrings(sharedStringsStream);

			ReadStyles(stylesStream);
		}

		private List<XlsbWorksheet> sheets;

		public List<XlsbWorksheet> Sheets
		{
			get { return sheets; }
			set { sheets = value; }
		}

		private XlsbSST _SST;

		public XlsbSST SST
		{
			get { return _SST; }
		}

		private XlsbStyles _Styles;

		public XlsbStyles Styles
		{
			get { return _Styles; }
		}


		private void ReadStyles(Stream fileStream)
		{
			throw new NotImplementedException();
		}

		private void ReadSharedStrings(Stream fileStream)
		{
			if (null == fileStream) return;

			throw new NotImplementedException();
		}

		private void ReadWorkbook(Stream fileStream)
		{
			//TODO: Try finaly, release resources

			byte[] buffer = new BinaryReader(fileStream).ReadBytes((int)fileStream.Length);

			UInt32 offset = 0;

			

			while (offset < buffer.Length)
			{
				XlsbRecord record = XlsbRecord.GetRecord(buffer, offset);
				short recId = 0;
				UInt32 reclen = 0;

				recId = (short)record.ID;
				reclen = (UInt32)record.GetLength();

				if (recId == 0 && reclen == 0)
					continue;


				//BaseRecord rec = (BaseRecord) Activator.CreateInstance(h[recid].GetType());

				//if (recHandler != null)
				//{
				//    Console.Write(String.Format("<{0}>\r\n[rec=0x{1:X} len=0x{2:X}]", recHandler.GetTag(), recid, reclen));

				//    for (int i = 0; i < reclen; i++)
				//    {
				//        Console.Write(String.Format(" {0:X2}", buffer[offset + i]));
				//    }

				//    Console.WriteLine();

				//    recHandler.Read(buffer, ref offset, (int)recid, (int)reclen, h, w);

				//    if (offset == UInt32.MaxValue)
				//    {
				//        Console.WriteLine("***Damaged buffer***");
				//        break;
				//    }

				//}
				//else
				//{
				//    Console.Write(String.Format("[rec=0x{0:X} len=0x{1:X}]", recid, reclen));

				//    for (int i = 0; i < reclen; i++)
				//    {
				//        Console.Write(String.Format(" {0:X2}", buffer[offset + i]));
				//    }

				//    Console.WriteLine();
				//}

				offset += reclen;

				//Console.WriteLine();
			}
		}
	}
}
