//#define DEBUGREADERS

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Xml;
using System.Globalization;
using ExcelDataReader.Core;
using ExcelDataReader.Core.OpenXmlFormat;

namespace Excel
{
	
	public class ExcelOpenXmlReader : IExcelDataReader
	{
		private const string ElementSheet = "sheet";
		private const string ElementT = "t";
		private const string ElementStringItem = "si";
		private const string ElementCellCrossReference = "cellXfs";
		private const string ElementNumberFormats = "numFmts";

		private const string AttributeSheetId = "sheetId";
		private const string AttributeVisibleState = "state";
		private const string AttributeName = "name";
		private const string AttributeRelationshipId = "r:id";

		private const string ElementRelationship = "Relationship";
		private const string AttributeId = "Id";
		private const string AttributeTarget = "Target";

	    #region Members

		private XlsxWorkbook m_workbook;

	    private bool m_isFirstRead;

	    private int m_resultIndex;
		private int m_emptyRowCount;
		private ZipWorker m_zipWorker;
		private XmlReader m_xmlReader;
		private Stream m_sheetStream;
		private string[] m_cellsNames;
		private object[] m_cellsValues;
		private object[] m_savedCellsValues;

	    private readonly List<int> m_defaultDateTimeStyles;
		private string m_namespaceUri;

		#endregion

		public ExcelOpenXmlReader(Stream stream)
		{
		    m_isFirstRead = true;

			m_defaultDateTimeStyles = new List<int>(new[] 
			{
				14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
			});

            m_zipWorker = new ZipWorker();
            m_zipWorker.Open(stream);

            ReadGlobals();
        }

		private void ReadGlobals()
		{
			List<XlsxWorksheet> sheets;
			XlsxSST sst;
			XlsxStyles styles;

			using (var stream = m_zipWorker.GetWorkbookStream()) {
				sheets = ReadWorkbook(stream);
			}
			using (var stream = m_zipWorker.GetWorkbookRelsStream()) {
				ReadWorkbookRels(stream, sheets);
			}

			using (var stream = m_zipWorker.GetSharedStringsStream()) {
				sst = ReadSharedStrings(stream);
			}

			using (var stream = m_zipWorker.GetStylesStream()) {
				styles = ReadStyles(stream);
			}

			m_workbook = new XlsxWorkbook(sheets, sst, styles);

			CheckDateTimeNumFmts(m_workbook.Styles.NumFmts);

		}
		private static List<XlsxWorksheet> ReadWorkbook(Stream xmlFileStream) {
			var sheets = new List<XlsxWorksheet>();

			using (XmlReader reader = XmlReader.Create(xmlFileStream)) {
				while (reader.Read()) {
					if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementSheet) {
						sheets.Add(new XlsxWorksheet(
											   reader.GetAttribute(AttributeName),
											   int.Parse(reader.GetAttribute(AttributeSheetId)),
											   reader.GetAttribute(AttributeRelationshipId),
											   reader.GetAttribute(AttributeVisibleState)));
					}

				}
			}
			return sheets;
		}

		private static void ReadWorkbookRels(Stream xmlFileStream, List<XlsxWorksheet> sheets) {
			using (XmlReader reader = XmlReader.Create(xmlFileStream)) {
				while (reader.Read()) {
					if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementRelationship) {
						string rid = reader.GetAttribute(AttributeId);

						for (int i = 0; i < sheets.Count; i++) {
							XlsxWorksheet tempSheet = sheets[i];

							if (tempSheet.RID == rid) {
								tempSheet.Path = reader.GetAttribute(AttributeTarget);
								sheets[i] = tempSheet;
								break;
							}
						}
					}

				}
			}
		}

		private static XlsxSST ReadSharedStrings(Stream xmlFileStream) {
			if (null == xmlFileStream)
				return null;

			var sst = new XlsxSST();

            using (XmlReader reader = XmlReader.Create(xmlFileStream))
            {
                // Skip phonetic string data.
                bool bSkipPhonetic = false;
                // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                bool bAddStringItem = false;
                string sStringItem = "";

                while (reader.Read())
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementStringItem)
                    {
                        // Do not add the string item until the next string item is read.
                        if (bAddStringItem)
                        {
                            // Add the string item to XlsxSST.
                            sst.Add(sStringItem);
                        }
                        else
                        {
                            // Add the string items from here on.
                            bAddStringItem = true;
                        }

                        // Reset the string item.
                        sStringItem = "";
                    }
                    else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementT)
                    {
                        // Skip phonetic string data.
                        if (!bSkipPhonetic)
                        {
                            // Append to the string item.
                            sStringItem += reader.ReadElementContentAsString();
                        }
                    }
                    if (reader.LocalName == "rPh")
                    {
                        // Phonetic items represents pronunciation hints for some East Asian languages.
                        // In the file 'xl/sharedStrings.xml', the phonetic properties appear like:
                        // <si>
                        //  <t>(a japanese text in KANJI)</t>
                        //  <rPh sb="0" eb="1">
                        //      <t>(its pronounciation in KATAKANA)</t>
                        //  </rPh>
                        // </si>
                        if (reader.NodeType == XmlNodeType.Element)
                            bSkipPhonetic = true;
                        else if (reader.NodeType == XmlNodeType.EndElement)
                            bSkipPhonetic = false;
                    }
                }
                // Do not add the last string item unless we have read previous string items.
                if (bAddStringItem)
                {
                    // Add the string item to XlsxSST.
                    sst.Add(sStringItem);
                }

            }
			return sst;
		}

		private static XlsxStyles ReadStyles(Stream xmlFileStream) {
			var styles = new XlsxStyles();

			if (null == xmlFileStream)
				return styles;

			bool rXlsxNumFmt = false;

			using (XmlReader reader = XmlReader.Create(xmlFileStream)) {
				while (reader.Read()) {
					if (!rXlsxNumFmt && reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementNumberFormats) {
						while (reader.Read()) {
							if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1)
								break;

							if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxNumFmt.N_numFmt) {
								styles.NumFmts.Add(
									new XlsxNumFmt(
										int.Parse(reader.GetAttribute(XlsxNumFmt.A_numFmtId)),
										reader.GetAttribute(XlsxNumFmt.A_formatCode)
										));
							}
						}

						rXlsxNumFmt = true;
					}

					if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementCellCrossReference) {
						while (reader.Read()) {
							if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1)
								break;

							if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxXf.N_xf) {
								var xfId = reader.GetAttribute(XlsxXf.A_xfId);
								var numFmtId = reader.GetAttribute(XlsxXf.A_numFmtId);

								styles.CellXfs.Add(
									new XlsxXf(
										xfId == null ? -1 : int.Parse(xfId),
										numFmtId == null ? -1 : int.Parse(numFmtId),
										reader.GetAttribute(XlsxXf.A_applyNumberFormat)
										));
							}
						}

						break;
					}
				}
			}
			return styles;
		}

		private void CheckDateTimeNumFmts(List<XlsxNumFmt> list)
		{
			if (list.Count == 0) return;

			foreach (XlsxNumFmt numFmt in list)
			{
				if (string.IsNullOrEmpty(numFmt.FormatCode)) continue;
				string fc = numFmt.FormatCode.ToLower();

				int pos;
				while ((pos = fc.IndexOf('"')) > 0)
				{
					int endPos = fc.IndexOf('"', pos + 1);

					if (endPos > 0) fc = fc.Remove(pos, endPos - pos + 1);
				}

				//it should only detect it as a date if it contains
				//dd mm mmm yy yyyy
				//h hh ss
				//AM PM
				//and only if these appear as "words" so either contained in [ ]
				//or delimted in someway
				//updated to not detect as date if format contains a #
				var formatReader = new FormatReader() {FormatString = fc};
				if (formatReader.IsDateFormatString())
				{
					m_defaultDateTimeStyles.Add(numFmt.Id);
				}
			}
		}

		private void ReadSheetGlobals(XlsxWorksheet sheet)
		{
		    ((IDisposable)m_xmlReader)?.Dispose();
		    m_sheetStream?.Dispose();

		    m_sheetStream = m_zipWorker.GetWorksheetStream(sheet.Path);

			if (null == m_sheetStream)
                return;

			m_xmlReader = XmlReader.Create(m_sheetStream);

			//count rows and cols in case there is no dimension elements
			int rows = 0;
			int cols = 0;

		    bool foundDimension = false;

			m_namespaceUri = null;
		    int biggestColumn = 0; //used when no col elements and no dimension
		    int cellElementsInRow = 0;
			while (m_xmlReader.Read())
			{
				if (m_xmlReader.NodeType == XmlNodeType.Element && m_xmlReader.LocalName == XlsxWorksheet.N_worksheet)
				{
					//grab the namespaceuri from the worksheet element
					m_namespaceUri = m_xmlReader.NamespaceURI;
				}
				
				if (m_xmlReader.NodeType == XmlNodeType.Element && m_xmlReader.LocalName == XlsxWorksheet.N_dimension)
				{
					string dimValue = m_xmlReader.GetAttribute(XlsxWorksheet.A_ref);

                    var dimension = new XlsxDimension(dimValue);
				    if (dimension.IsRange)
				    {
				        sheet.Dimension = dimension;
				        foundDimension = true;

                        break;
				    }
				}

                //removed: Do not use col to work out number of columns as this is really for defining formatting, so may not contain all columns
                //if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_col)
                //    cols++;

			    if (m_xmlReader.NodeType == XmlNodeType.Element && m_xmlReader.LocalName == XlsxWorksheet.N_row)
			    {
			        biggestColumn = Math.Max(biggestColumn, cellElementsInRow);
			        cellElementsInRow = 0;
                    rows++;
			    }

			    //check cells so we can find size of sheet if can't work it out from dimension or col elements (dimension should have been set before the cells if it was available)
                //ditto for cols
                if (cols == 0 && m_xmlReader.NodeType == XmlNodeType.Element && m_xmlReader.LocalName == XlsxWorksheet.N_c)
                {
                    cellElementsInRow++; 

                    var refAttribute = m_xmlReader.GetAttribute(XlsxWorksheet.A_r);

                    if (refAttribute != null)
                    {
                        int column;
                        ReferenceHelper.ParseReference(refAttribute, out column);
                        if (column > biggestColumn)
                            biggestColumn = column;
                    }
                }
			}

            biggestColumn = Math.Max(biggestColumn, cellElementsInRow);

            //if we didn't get a dimension element then use the calculated rows/cols to create it
            if (!foundDimension)
			{
                if (cols == 0)
                    cols = biggestColumn;

				if (rows == 0 || cols == 0) 
				{
					sheet.IsEmpty = true;
					return;
				}

				sheet.Dimension = new XlsxDimension(rows, cols);

				//we need to reset our position to sheet data
				((IDisposable)m_xmlReader).Dispose();
                m_sheetStream.Dispose();
                m_sheetStream = m_zipWorker.GetWorksheetStream(sheet.Path);
				m_xmlReader = XmlReader.Create(m_sheetStream);

			}

			//read up to the sheetData element. if this element is empty then there aren't any rows and we need to null out dimension

			m_xmlReader.ReadToFollowing(XlsxWorksheet.N_sheetData, m_namespaceUri);
			if (m_xmlReader.IsEmptyElement)
			{
				sheet.IsEmpty = true;
			}
		}

		private bool ReadSheetRow(XlsxWorksheet sheet)
		{
			if (null == m_xmlReader)
                return false;

			if (m_emptyRowCount != 0)
			{
				m_cellsValues = new object[sheet.ColumnsCount];
				m_emptyRowCount--;
				Depth++;

				return true;
			}

			if (m_savedCellsValues != null)
			{
				m_cellsValues = m_savedCellsValues;
				m_savedCellsValues = null;
				Depth++;

				return true;
			}

            if ((m_xmlReader.NodeType == XmlNodeType.Element && m_xmlReader.LocalName == XlsxWorksheet.N_row) ||
                m_xmlReader.ReadToFollowing(XlsxWorksheet.N_row, m_namespaceUri))
			{
				m_cellsValues = new object[sheet.ColumnsCount];

                int rowIndex;
                if (!int.TryParse(m_xmlReader.GetAttribute(XlsxWorksheet.A_r), out rowIndex))
                    rowIndex = Depth + 1;

				if (rowIndex != (Depth + 1))
				{
					m_emptyRowCount = rowIndex - Depth - 1;
				}
				bool hasValue = false;
				string a_s = string.Empty;
				string a_t = string.Empty;
				string a_r = string.Empty;
				int col = 0;

			    while (m_xmlReader.Read())
                {
                    if (m_xmlReader.Depth == 2) break;

                    if (m_xmlReader.NodeType == XmlNodeType.Element)
                    {
                        hasValue = false;

                        if (m_xmlReader.LocalName == XlsxWorksheet.N_c)
                        {
                            a_s = m_xmlReader.GetAttribute(XlsxWorksheet.A_s);
                            a_t = m_xmlReader.GetAttribute(XlsxWorksheet.A_t);
                            a_r = m_xmlReader.GetAttribute(XlsxWorksheet.A_r);

                            if (a_r != null)
                            {
                                ReferenceHelper.ParseReference(a_r, out col);
                            }
                            else
                                ++col;
                        }
                        else if (m_xmlReader.LocalName == XlsxWorksheet.N_v || m_xmlReader.LocalName == XlsxWorksheet.N_t)
                        {
                            hasValue = true;
                        }
                    }

                    if (m_xmlReader.NodeType == XmlNodeType.Text && hasValue)
                    {
                    	double number;
                        object o = m_xmlReader.Value;

	                    const NumberStyles style = NumberStyles.Any;
						var culture = CultureInfo.InvariantCulture;
                        
                        if (double.TryParse(o.ToString(), style, culture, out number))
                            o = number;
                        	
                        if (null != a_t && a_t == XlsxWorksheet.A_s) //if string
                        {
                            o = Helpers.ConvertEscapeChars(m_workbook.SST[int.Parse(o.ToString())]);
                        } // Requested change 4: missing (it appears that if should be else if)
                        else if (null != a_t && a_t == XlsxWorksheet.N_inlineStr) //if string inline
                        {
                            o = Helpers.ConvertEscapeChars(o.ToString());
                        }
                        else if (a_t == "b") //boolean
						{
							o = m_xmlReader.Value == "1";
						}  
                        else if (null != a_s) //if something else
                        {
                            XlsxXf xf = m_workbook.Styles.CellXfs[int.Parse(a_s)];
                            if (o != null && o.ToString() != string.Empty && IsDateTimeStyle(xf.NumFmtId))
                                o = Helpers.ConvertFromOATime(number);
                            else if (xf.NumFmtId == 49)
                                o = o.ToString();
                        }
                                                


                        if (col - 1 < m_cellsValues.Length)
                            m_cellsValues[col - 1] = o;
                    }
                }

				if (m_emptyRowCount > 0)
				{
					m_savedCellsValues = m_cellsValues;
					return ReadSheetRow(sheet);
				}
				Depth++;

				return true;
			}

			((IDisposable)m_xmlReader).Dispose();
		    m_sheetStream?.Dispose();

		    return false;
		}

		private bool InitializeSheetRead()
		{
			if (ResultsCount <= 0) return false;

			ReadSheetGlobals(m_workbook.Sheets[m_resultIndex]);

			if (m_workbook.Sheets[m_resultIndex].Dimension == null) return false;

			m_isFirstRead = false;

			Depth = 0;
			m_emptyRowCount = 0;

			return true;
		}

		private bool IsDateTimeStyle(int styleId)
		{
			return m_defaultDateTimeStyles.Contains(styleId);
		}


		#region IExcelDataReader Members

		public void Reset() {
			m_resultIndex = 0;
			m_isFirstRead = true;
			m_savedCellsValues = null;
		}

		public bool IsFirstRowAsColumnNames { get; set; }

	    public bool ConvertOaDate { get; set; }

	    public ReadOption ReadOption { get; set; }

	    public Encoding Encoding => null;

	    public string Name => m_resultIndex >= 0 && m_resultIndex < ResultsCount ? m_workbook.Sheets[m_resultIndex].Name : null;

	    public string VisibleState => m_resultIndex >= 0 && m_resultIndex < ResultsCount ? m_workbook.Sheets[m_resultIndex].VisibleState : null;

	    public void Close()
		{
            if (IsClosed)
                return;

		    ((IDisposable)m_xmlReader)?.Dispose();
		    m_sheetStream?.Dispose();
		    m_zipWorker?.Dispose();

            m_zipWorker = null;
            m_xmlReader = null;
            m_sheetStream = null;

            m_workbook = null;
            m_cellsValues = null;
            m_savedCellsValues = null;

            IsClosed = true;
        }

        public int Depth { get; private set; }

	    public int ResultsCount => m_workbook?.Sheets.Count ?? -1;

	    public bool IsClosed { get; private set; }

	    public bool NextResult()
		{
			if (m_resultIndex >= ResultsCount - 1)
                return false;

			m_resultIndex++;

			m_isFirstRead = true;
		    m_savedCellsValues = null;

			return true;
		}

        public bool Read()
        {
            if (m_isFirstRead)
            {
                var initializeSheetRead = InitializeSheetRead();
                if (!initializeSheetRead)
                    return false;

				if (IsFirstRowAsColumnNames)
                {
					if (ReadSheetRow(m_workbook.Sheets[m_resultIndex]))
                    {
						m_cellsNames = new string[m_cellsValues.Length];
						for (var i = 0; i < m_cellsValues.Length; i++)
                        {
							var value = m_cellsValues[i]?.ToString();
							if (!string.IsNullOrEmpty(value))
								m_cellsNames[i] = value;
						}
					}
                    else
                    {
						return false;
					}
				}
                else
                {
					m_cellsNames = null;
				}

			}

            return ReadSheetRow(m_workbook.Sheets[m_resultIndex]);
        }

		public int FieldCount => m_resultIndex >= 0 && m_resultIndex < ResultsCount ? m_workbook.Sheets[m_resultIndex].ColumnsCount : -1;

	    public bool GetBoolean(int i)
		{
		    return !IsDBNull(i) && bool.Parse(m_cellsValues[i].ToString());
		}

		public DateTime GetDateTime(int i)
		{
			if (IsDBNull(i)) return DateTime.MinValue;

			try
			{
				return (DateTime)m_cellsValues[i];
			}
			catch (InvalidCastException)
			{
				return DateTime.MinValue;
			}
		}

		public decimal GetDecimal(int i)
		{
		    return IsDBNull(i) ? decimal.MinValue : decimal.Parse(m_cellsValues[i].ToString());
		}

		public double GetDouble(int i)
		{
		    return IsDBNull(i) ? double.MinValue : double.Parse(m_cellsValues[i].ToString());
		}

		public float GetFloat(int i)
		{
		    return IsDBNull(i) ? float.MinValue : float.Parse(m_cellsValues[i].ToString());
		}

		public short GetInt16(int i)
		{
		    return IsDBNull(i) ? short.MinValue : short.Parse(m_cellsValues[i].ToString());
		}

		public int GetInt32(int i)
		{
		    return IsDBNull(i) ? int.MinValue : int.Parse(m_cellsValues[i].ToString());
		}

		public long GetInt64(int i)
		{
		    return IsDBNull(i) ? long.MinValue : long.Parse(m_cellsValues[i].ToString());
		}

		public string GetString(int i)
		{
		    return IsDBNull(i) ? null : m_cellsValues[i].ToString();
		}

		public object GetValue(int i)
		{
			return m_cellsValues[i];
		}

		public bool IsDBNull(int i)
		{
            return null == m_cellsValues[i];
		}

		public object this[int i] => m_cellsValues[i];

	    #endregion

		#region IDisposable Members

		public void Dispose()
		{
			Dispose(true);

			GC.SuppressFinalize(this);
		}

		private void Dispose(bool disposing)
		{
		    if (disposing)
		        Close();
		}

		~ExcelOpenXmlReader()
		{
			Dispose(false);
		}

		#endregion

		#region  Not Supported IDataReader Members


		public int RecordsAffected
		{
			get { throw new NotSupportedException(); }
		}

		#endregion

		#region Not Supported IDataRecord Members


		public byte GetByte(int i)
		{
			throw new NotSupportedException();
		}

		public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
		{
			throw new NotSupportedException();
		}

		public char GetChar(int i)
		{
			throw new NotSupportedException();
		}

		public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
		{
			throw new NotSupportedException();
		}

		public IDataReader GetData(int i)
		{
			throw new NotSupportedException();
		}

		public string GetDataTypeName(int i)
		{
			throw new NotSupportedException();
		}

		public Type GetFieldType(int i)
		{
			if (m_cellsValues[i] == null)
				return null;
			return m_cellsValues[i].GetType();
		}

		public Guid GetGuid(int i)
		{
			throw new NotSupportedException();
		}

		public string GetName(int i)
		{
		    return m_cellsNames?[i];
		}

		public int GetOrdinal(string name)
		{
			throw new NotSupportedException();
		}

		public int GetValues(object[] values)
		{
			throw new NotSupportedException();
		}

		public object this[string name]
		{
			get { throw new NotSupportedException(); }
		}

	    /// <inheritdoc />
	    public DataTable GetSchemaTable()
	    {
	        throw new NotSupportedException();
	    }

	    #endregion


	}
}
