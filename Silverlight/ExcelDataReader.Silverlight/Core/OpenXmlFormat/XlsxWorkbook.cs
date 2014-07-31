namespace ExcelDataReader.Silverlight.Core.OpenXmlFormat
{
	using System;
	using System.Collections.Generic;
	using System.Xml;
	using System.IO;

	internal class XlsxWorkbook
	{
		private const string N_sheet = "sheet";
		private const string N_t = "t";
        private const string N_si = "si";
		private const string N_cellXfs = "cellXfs";
		private const string N_numFmts = "numFmts";

		private const string A_sheetId = "sheetId";
		private const string A_name = "name";
		private const string A_rid = "r:id";

		private const string N_rel = "Relationship";
		private const string A_id = "Id";
		private const string A_target = "Target";

		private XlsxWorkbook() { }

        public XlsxWorkbook(byte[] workbookByteArray, byte[] relsByteArray, byte[] sharedStringsByteArray, byte[] stylesByteArray)
		{
            if (null == workbookByteArray) throw new ArgumentNullException();

            ReadWorkbook(workbookByteArray);

            ReadWorkbookRels(relsByteArray);

            ReadSharedStrings(sharedStringsByteArray);
            
            ReadStyles(stylesByteArray);
		}

		private List<XlsxWorksheet> sheets;

		public List<XlsxWorksheet> Sheets
		{
			get { return sheets; }
			set { sheets = value; }
		}

		private XlsxSST _SST;

		public XlsxSST SST
		{
			get { return _SST; }
		}

		private XlsxStyles _Styles;

		public XlsxStyles Styles
		{
			get { return _Styles; }
		}


        private void ReadStyles(byte[] byteArray)
		{
            if (byteArray == null) return;

			_Styles = new XlsxStyles();

			bool rXlsxNumFmt = false;

            Stream xmlFileStream = new MemoryStream(byteArray);

			using (XmlReader reader = XmlReader.Create(xmlFileStream))
			{
				while (reader.Read())
				{
					if (!rXlsxNumFmt && reader.NodeType == XmlNodeType.Element && reader.Name == N_numFmts)
					{
						while (reader.Read())
						{
							if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1) break;

							if (reader.NodeType == XmlNodeType.Element && reader.Name == XlsxNumFmt.N_numFmt)
							{
								_Styles.NumFmts.Add(
									new XlsxNumFmt(
										int.Parse(reader.GetAttribute(XlsxNumFmt.A_numFmtId)),
										reader.GetAttribute(XlsxNumFmt.A_formatCode)
										));
							}
						}

						rXlsxNumFmt = true;
					}

					if (reader.NodeType == XmlNodeType.Element && reader.Name == N_cellXfs)
					{
						while (reader.Read())
						{
							if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1) break;

							if (reader.NodeType == XmlNodeType.Element && reader.Name == XlsxXf.N_xf)
							{
								_Styles.CellXfs.Add(
									new XlsxXf(
										int.Parse(reader.GetAttribute(XlsxXf.A_xfId)),
										int.Parse(reader.GetAttribute(XlsxXf.A_numFmtId)),
										reader.GetAttribute(XlsxXf.A_applyNumberFormat)
										));
							}
						}

						break;
					}
				}

				xmlFileStream.Close();
			}
		}

        private void ReadSharedStrings(byte[] byteArray)
        {
            if (null == byteArray) return;

            _SST = new XlsxSST();

            Stream xmlFileStream = new MemoryStream(byteArray);

            using (XmlReader reader = XmlReader.Create(xmlFileStream))
            {
                // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                bool bAddStringItem = false;
                string sStringItem = "";

                while (reader.Read())
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    if (reader.NodeType == XmlNodeType.Element && reader.Name == N_si)
                    {
                        // Do not add the string item until the next string item is read.
                        if (bAddStringItem)
                        {
                            // Add the string item to XlsxSST.
                            _SST.Add(sStringItem);
                        }
                        else
                        {
                            // Add the string items from here on.
                            bAddStringItem = true;
                        }

                        // Reset the string item.
                        sStringItem = "";
                    }

                    if (reader.NodeType == XmlNodeType.Element && reader.Name == N_t)
                    {
                        // Append to the string item.
                        sStringItem += reader.ReadElementContentAsString();
                    }
                }
                // Do not add the last string item unless we have read previous string items.
                if (bAddStringItem)
                {
                    // Add the string item to XlsxSST.
                    _SST.Add(sStringItem);
                }

                xmlFileStream.Close();
            }
        }

		private void ReadWorkbook(Byte[] byteArray)
		{
			sheets = new List<XlsxWorksheet>();

            Stream xmlFileStream = new MemoryStream(byteArray);

			using (XmlReader reader = XmlReader.Create(xmlFileStream))
			{
				while (reader.Read())
				{
					if (reader.NodeType == XmlNodeType.Element && reader.Name == N_sheet)
					{
						sheets.Add(new XlsxWorksheet(
											   reader.GetAttribute(A_name),
											   int.Parse(reader.GetAttribute(A_sheetId)), reader.GetAttribute(A_rid)));
					}

				}

				xmlFileStream.Close();
			}

		}

        private void ReadWorkbookRels(byte[] byteArray)
		{
            Stream xmlFileStream = new MemoryStream(byteArray);

			using (XmlReader reader = XmlReader.Create(xmlFileStream))
			{
				while (reader.Read())
				{
					if (reader.NodeType == XmlNodeType.Element && reader.Name == N_rel)
					{
						string rid = reader.GetAttribute(A_id);

						for (int i = 0; i < sheets.Count; i++)
						{
							XlsxWorksheet tempSheet = sheets[i];

							if (tempSheet.RID == rid)
							{
								tempSheet.Path = reader.GetAttribute(A_target);
								sheets[i] = tempSheet;
								break;
							}
						}
					}

				}

				xmlFileStream.Close();
			}
		}
	}
}