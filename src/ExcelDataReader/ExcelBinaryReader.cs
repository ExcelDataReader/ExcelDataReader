using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Core;
using ExcelDataReader.Core.BinaryFormat;
using ExcelDataReader.Data;
using ExcelDataReader.Exceptions;
using ExcelDataReader.Log;
using ExcelDataReader.Misc;

namespace Excel
{
	/// <summary>
	/// ExcelDataReader Class
	/// </summary>
	public class ExcelBinaryReader : IExcelDataReader
	{
	    #region Members

		private Stream m_file;
		private XlsHeader m_hdr;
		private List<XlsWorksheet> m_sheets;
		private XlsBiffStream m_stream;
        //private DataSet m_workbookData;
		private XlsWorkbookGlobals m_globals;
		private ushort m_version;
        private Encoding m_encoding;
		private bool m_isValid;
		private bool m_isClosed;
        private readonly Encoding m_Default_Encoding = Encoding.Unicode;
		private string m_exceptionMessage;
		private string[] m_cellsNames;
		private object[] m_cellsValues;
		private uint[] m_dbCellAddrs;
		private int m_dbCellAddrsIndex;
		private int m_SheetIndex;
		private int m_depth;
		private int m_cellOffset;
		private int m_maxCol;
		private int m_maxRow;
		private bool m_noIndex;
		private XlsBiffRow m_currentRowRecord;
		private ReadOption readOption = ReadOption.Strict;
		private bool m_IsFirstRead;

	    private const string WORKBOOK = "Workbook";
		private const string BOOK = "Book";
        private const string COLUMN = "Column";

		private bool disposed;

	    private bool convertOaDate = true;

	    #endregion

		public ExcelBinaryReader()
		{
		    m_encoding = m_Default_Encoding;
			m_version = 0x0600;
			m_isValid = true;
			m_SheetIndex = -1;
			m_IsFirstRead = true;
		}

		#region IDisposable Members

		public void Dispose()
		{
			Dispose(true);

			GC.SuppressFinalize(this);
		}

		private void Dispose(bool disposing)
		{
			// Check to see if Dispose has already been called.
			if (!this.disposed)
			{
				if (disposing)
				{
                    //if (m_workbookData != null) m_workbookData.Dispose();

					if (m_sheets != null) m_sheets.Clear();
				}

                //m_workbookData = null;
				m_sheets = null;
				m_stream = null;
				m_globals = null;
				m_encoding = null;
				m_hdr = null;

				disposed = true;
			}
		}

		~ExcelBinaryReader()
		{
			Dispose(false);
		}

		#endregion

		#region Private methods

		private int findFirstDataCellOffset(int startOffset)
		{
			//seek to the first dbcell record
			var record = m_stream.ReadAt(startOffset);
			while (!(record is XlsBiffDbCell))
			{
				if (m_stream.Position >= m_stream.Size)
					return -1;

				if (record is XlsBiffEOF)
					return -1;

				record = m_stream.Read();
			}

			XlsBiffDbCell startCell = (XlsBiffDbCell)record;
			XlsBiffRow row = null;

			int offs = startCell.RowAddress;

			do
			{
				row = m_stream.ReadAt(offs) as XlsBiffRow;
				if (row == null) break;

				offs += row.Size;

			} while (null != row);

			return offs;
		}

		private void readWorkBookGlobals()
		{
			//Read Header
			try
			{
				m_hdr = XlsHeader.ReadHeader(m_file);
			}
			catch (HeaderException ex)
			{
				fail(ex.Message);
				return;
			}
			catch (FormatException ex)
			{
				fail(ex.Message);
				return;
			}

			XlsRootDirectory dir = new XlsRootDirectory(m_hdr);
			XlsDirectoryEntry workbookEntry = dir.FindEntry(WORKBOOK) ?? dir.FindEntry(BOOK);

			if (workbookEntry == null)
			{ fail(Errors.ErrorStreamWorkbookNotFound); return; }

			if (workbookEntry.EntryType != STGTY.STGTY_STREAM)
			{ fail(Errors.ErrorWorkbookIsNotStream); return; }

			m_stream = new XlsBiffStream(m_hdr, workbookEntry.StreamFirstSector, workbookEntry.IsEntryMiniStream, dir, this);

			m_globals = new XlsWorkbookGlobals();

			m_stream.Seek(0, SeekOrigin.Begin);

			XlsBiffRecord rec = m_stream.Read();
			XlsBiffBOF bof = rec as XlsBiffBOF;

			if (bof == null || bof.Type != BIFFTYPE.WorkbookGlobals)
			{ fail(Errors.ErrorWorkbookGlobalsInvalidData); return; }

			bool sst = false;

			m_version = bof.Version;
			m_sheets = new List<XlsWorksheet>();

			while (null != (rec = m_stream.Read()))
			{
				switch (rec.ID)
				{
					case BIFFRECORDTYPE.INTERFACEHDR:
						m_globals.InterfaceHdr = (XlsBiffInterfaceHdr)rec;
						break;
					case BIFFRECORDTYPE.BOUNDSHEET:
						XlsBiffBoundSheet sheet = (XlsBiffBoundSheet)rec;

						if (sheet.Type != XlsBiffBoundSheet.SheetType.Worksheet) break;

						sheet.IsV8 = isV8();
						//sheet.UseEncoding = Encoding;
						LogManager.Log(this).Debug("BOUNDSHEET IsV8={0}", sheet.IsV8);

						m_sheets.Add(new XlsWorksheet(m_globals.Sheets.Count, sheet));
						m_globals.Sheets.Add(sheet);

						break;
					case BIFFRECORDTYPE.MMS:
						m_globals.MMS = rec;
						break;
					case BIFFRECORDTYPE.COUNTRY:
						m_globals.Country = rec;
						break;
					case BIFFRECORDTYPE.CODEPAGE:

						m_globals.CodePage = (XlsBiffSimpleValueRecord)rec;

                        //set encoding based on code page name
                        //PCL does not supported codepage numbers
                        if (m_globals.CodePage.Value == 1200)
                            m_encoding = EncodingHelper.GetEncoding(65001);
                        else
                            m_encoding = EncodingHelper.GetEncoding(m_globals.CodePage.Value);
                        //note: the format spec states that for BIFF8 this is always UTF-16.
                        //as PCL does not supported codepage numbers, it is best to assume UTF-16 for encoding
                        //try
                        //{
                        //    m_encoding = Encoding.GetEncoding(m_globals.CodePage.Value);
                        //}
                        //catch (ArgumentException)
                        //{
                        //    // Warning - Password protection
                        //}

						break;
					case BIFFRECORDTYPE.FONT:
					case BIFFRECORDTYPE.FONT_V34:
						m_globals.Fonts.Add(rec);
						break;
                    case BIFFRECORDTYPE.FORMAT_V23:
				        {
				            var fmt = (XlsBiffFormatString) rec;
                            //fmt.UseEncoding = m_encoding;
				            m_globals.Formats.Add((ushort) m_globals.Formats.Count, fmt);
				        }
                        break;
                    case BIFFRECORDTYPE.FORMAT:
				        {
				            var fmt = (XlsBiffFormatString) rec;
                            m_globals.Formats.Add(fmt.Index, fmt);
				        }
				        break;
					case BIFFRECORDTYPE.XF:
					case BIFFRECORDTYPE.XF_V4:
					case BIFFRECORDTYPE.XF_V3:
					case BIFFRECORDTYPE.XF_V2:
						m_globals.ExtendedFormats.Add(rec);
						break;
					case BIFFRECORDTYPE.SST:
						m_globals.SST = (XlsBiffSST)rec;
						sst = true;
						break;
					case BIFFRECORDTYPE.CONTINUE:
						if (!sst) break;
						XlsBiffContinue contSST = (XlsBiffContinue)rec;
						m_globals.SST.Append(contSST);
						break;
					case BIFFRECORDTYPE.EXTSST:
						m_globals.ExtSST = rec;
						sst = false;
						break;
					case BIFFRECORDTYPE.PROTECT:
					case BIFFRECORDTYPE.PASSWORD:
					case BIFFRECORDTYPE.PROT4REVPASSWORD:
						//IsProtected
						break;
					case BIFFRECORDTYPE.EOF:
						if (m_globals.SST != null)
							m_globals.SST.ReadStrings();
						return;

					default:
						continue;
				}
			}
		}

		private bool readWorkSheetGlobals(XlsWorksheet sheet, out XlsBiffIndex idx, out XlsBiffRow row)
		{
			idx = null;
			row = null;

			m_stream.Seek((int)sheet.DataOffset, SeekOrigin.Begin);

			XlsBiffBOF bof = m_stream.Read() as XlsBiffBOF;
			if (bof == null || bof.Type != BIFFTYPE.Worksheet) return false;

			//DumpBiffRecords();

			XlsBiffRecord rec = m_stream.Read();
			if (rec == null) return false;
			if (rec is XlsBiffIndex)
			{
				idx = rec as XlsBiffIndex;
			}
			else if (rec is XlsBiffUncalced)
			{
				// Sometimes this come before the index...
				idx = m_stream.Read() as XlsBiffIndex;
			}

			//if (null == idx)
			//{
			//	// There is a record before the index! Chech his type and see the MS Biff Documentation
			//	return false;
			//}

			if (idx != null)
			{
				idx.IsV8 = isV8();
				LogManager.Log(this).Debug("INDEX IsV8={0}", idx.IsV8);
			}
				


			XlsBiffRecord trec;
			XlsBiffDimensions dims = null;

			do
			{
				trec = m_stream.Read();
				if (trec.ID == BIFFRECORDTYPE.DIMENSIONS)
				{
					dims = (XlsBiffDimensions)trec;
					break;
				}

			} while (trec != null && trec.ID != BIFFRECORDTYPE.ROW);

			//if we are already on row record then set that as the row, otherwise step forward till we get to a row record
			if (trec.ID == BIFFRECORDTYPE.ROW)
				row = (XlsBiffRow)trec;

			XlsBiffRow rowRecord = null;
			while (rowRecord == null)
			{
				if (m_stream.Position >= m_stream.Size)
					break;
				var thisRec = m_stream.Read();

				LogManager.Log(this).Debug("finding rowRecord offset {0}, rec: {1}", thisRec.Offset, thisRec.ID);
				if (thisRec is XlsBiffEOF)
					break;
				rowRecord = thisRec as XlsBiffRow;
			}

			if (rowRecord != null)
				LogManager.Log(this).Debug("Got row {0}, rec: id={1},rowindex={2}, rowColumnStart={3}, rowColumnEnd={4}", rowRecord.Offset, rowRecord.ID, rowRecord.RowIndex, rowRecord.FirstDefinedColumn, rowRecord.LastDefinedColumn);

			row = rowRecord;

            if (dims != null) {
                dims.IsV8 = isV8();
				LogManager.Log(this).Debug("dims IsV8={0}", dims.IsV8);
                m_maxCol = dims.LastColumn - 1;

				//handle case where sheet reports last column is 1 but there are actually more
				if (m_maxCol <= 0 && rowRecord != null)
				{
					m_maxCol = rowRecord.LastDefinedColumn;
				}
				
                m_maxRow = (int)dims.LastRow;
                sheet.Dimensions = dims;
            } else {
                m_maxCol = 256;
                m_maxRow = (int)idx.LastExistingRow;
            }

			if (idx != null && idx.LastExistingRow <= idx.FirstExistingRow)
			{
				return false;
			}
			else if (row == null)
			{
				return false;
			}

			m_depth = 0;

			return true;
		}

		private void DumpBiffRecords()
		{
			XlsBiffRecord rec = null;
			var startPos = m_stream.Position;

			do
			{
				rec = m_stream.Read();
				LogManager.Log(this).Debug(rec.ID.ToString());
			} while (rec != null && m_stream.Position < m_stream.Size);

			m_stream.Seek(startPos, SeekOrigin.Begin);
		}

		/// <returns>true if row was read successfully</returns>
		private bool readWorkSheetRow()
		{
			m_cellsValues = new object[m_maxCol];

			while (m_cellOffset < m_stream.Size)
			{
				XlsBiffRecord rec = m_stream.ReadAt(m_cellOffset);
				m_cellOffset += rec.Size;

				if ((rec is XlsBiffDbCell) || (rec is XlsBiffMSODrawing)) { break; };
				if (rec is XlsBiffEOF) { return false; };

				XlsBiffBlankCell cell = rec as XlsBiffBlankCell;

				if ((null == cell) || (cell.ColumnIndex >= m_maxCol)) continue;
				if (cell.RowIndex != m_depth) { m_cellOffset -= rec.Size; break; };

				pushCellValue(cell);
			}

			m_depth++;

			return m_cellOffset <= m_stream.Size;
		}

		
		
		private void pushCellValue(XlsBiffBlankCell cell)
		{
			double _dValue;
			LogManager.Log(this).Debug("pushCellValue {0}", cell.ID);
			switch (cell.ID)
			{
                case BIFFRECORDTYPE.BOOLERR:
                    if (cell.ReadByte(7) == 0)
                        m_cellsValues[cell.ColumnIndex] = cell.ReadByte(6) != 0;
                    break;
                case BIFFRECORDTYPE.BOOLERR_OLD:
                    if (cell.ReadByte(8) == 0)
                        m_cellsValues[cell.ColumnIndex] = cell.ReadByte(7) != 0;
                    break;
				case BIFFRECORDTYPE.INTEGER:
				case BIFFRECORDTYPE.INTEGER_OLD:
					m_cellsValues[cell.ColumnIndex] = ((XlsBiffIntegerCell)cell).Value;
					break;
				case BIFFRECORDTYPE.NUMBER:
				case BIFFRECORDTYPE.NUMBER_OLD:

					_dValue = ((XlsBiffNumberCell)cell).Value;

					m_cellsValues[cell.ColumnIndex] = !ConvertOaDate ?
						_dValue : tryConvertOADateTime(_dValue, cell.XFormat);

					LogManager.Log(this).Debug("VALUE: {0}", _dValue);
					break;
				case BIFFRECORDTYPE.LABEL:
				case BIFFRECORDTYPE.LABEL_OLD:
				case BIFFRECORDTYPE.RSTRING:
					
					m_cellsValues[cell.ColumnIndex] = ((XlsBiffLabelCell)cell).Value;

					LogManager.Log(this).Debug("VALUE: {0}", m_cellsValues[cell.ColumnIndex]);
					break;
				case BIFFRECORDTYPE.LABELSST:
					string tmp = m_globals.SST.GetString(((XlsBiffLabelSSTCell)cell).SSTIndex);
					LogManager.Log(this).Debug("VALUE: {0}", tmp);
					m_cellsValues[cell.ColumnIndex] = tmp;
					break;
				case BIFFRECORDTYPE.RK:

					_dValue = ((XlsBiffRKCell)cell).Value;

					m_cellsValues[cell.ColumnIndex] = !ConvertOaDate ?
						_dValue : tryConvertOADateTime(_dValue, cell.XFormat);

					LogManager.Log(this).Debug("VALUE: {0}", _dValue);
					break;
				case BIFFRECORDTYPE.MULRK:

					XlsBiffMulRKCell _rkCell = (XlsBiffMulRKCell)cell;
					for (ushort j = cell.ColumnIndex; j <= _rkCell.LastColumnIndex; j++)
					{
                        _dValue = _rkCell.GetValue(j);
						LogManager.Log(this).Debug("VALUE[{1}]: {0}", _dValue, j);
                        m_cellsValues[j] = !ConvertOaDate ? _dValue : tryConvertOADateTime(_dValue, _rkCell.GetXF(j));
					}

					break;
				case BIFFRECORDTYPE.BLANK:
				case BIFFRECORDTYPE.BLANK_OLD:
				case BIFFRECORDTYPE.MULBLANK:
					// Skip blank cells

					break;
				case BIFFRECORDTYPE.FORMULA:
				case BIFFRECORDTYPE.FORMULA_OLD:

					object _oValue = ((XlsBiffFormulaCell)cell).Value;

					if (null != _oValue && _oValue is FORMULAERROR)
					{
						_oValue = null;
					}
					else
					{
						m_cellsValues[cell.ColumnIndex] = !ConvertOaDate ?
							_oValue : tryConvertOADateTime(_oValue, (ushort)(cell.XFormat));//date time offset
					}

					LogManager.Log(this).Debug("VALUE: {0}", _oValue);
					break;
				default:
					break;
			}
		}

		private bool moveToNextRecord()
		{
			//if sheet has no index
			if (m_noIndex)
			{
				LogManager.Log(this).Debug("No index");
				return moveToNextRecordNoIndex();
			}

			//if sheet has index
			if (null == m_dbCellAddrs ||
				m_dbCellAddrsIndex == m_dbCellAddrs.Length ||
				m_depth == m_maxRow) return false;

			bool success = readWorkSheetRow();

			if (!success && m_dbCellAddrsIndex < (m_dbCellAddrs.Length - 1))
			{
				m_dbCellAddrsIndex++;
				m_cellOffset = findFirstDataCellOffset((int)m_dbCellAddrs[m_dbCellAddrsIndex]);
				if (m_cellOffset < 0)
					return false;
				success = readWorkSheetRow();
			}

			return success;
		}

		private bool moveToNextRecordNoIndex()
		{
			//seek from current row record to start of cell data where that cell relates to the next row record
			XlsBiffRow rowRecord = m_currentRowRecord;

			if (rowRecord == null)
				return false;
			
			if (rowRecord.RowIndex < m_depth)
			{
				m_stream.Seek(rowRecord.Offset + rowRecord.Size, SeekOrigin.Begin);
				do
				{
					if (m_stream.Position >= m_stream.Size)
						return false;

					var record = m_stream.Read();
					if (record is XlsBiffEOF)
						return false;

					rowRecord = record as XlsBiffRow;

				} while (rowRecord == null || rowRecord.RowIndex < m_depth);
			}

			m_currentRowRecord = rowRecord;
			//m_depth = m_currentRowRecord.RowIndex;

			//we have now found the row record for the new row, the we need to seek forward to the first cell record
			XlsBiffBlankCell cell = null;
			do
			{
				if (m_stream.Position >= m_stream.Size)
					return false;

				var record = m_stream.Read();
				if (record is XlsBiffEOF)
					return false;

				if (record.IsCell)
				{
					var candidateCell = record as XlsBiffBlankCell;
					if (candidateCell != null)
					{
						if (candidateCell.RowIndex == m_currentRowRecord.RowIndex)
							cell = candidateCell;
					}
				}
			} while (cell == null);

			m_cellOffset = cell.Offset;
			bool success = readWorkSheetRow();


			//read last row
			//if (!m_canRead && m_depth > 0) m_canRead = true;

			//if (!m_canRead && m_dbCellAddrsIndex < (m_dbCellAddrs.Length - 1))
			//{
			//	m_dbCellAddrsIndex++;
			//	m_cellOffset = findFirstDataCellOffset((int)m_dbCellAddrs[m_dbCellAddrsIndex]);

			//	m_canRead = readWorkSheetRow();
			//}

			return success;
		}

		private void initializeSheetRead()
		{
			if (m_SheetIndex == ResultsCount) return;

			m_dbCellAddrs = null;

			m_IsFirstRead = false;

			if (m_SheetIndex == -1) m_SheetIndex = 0;

			XlsBiffIndex idx;

			if (!readWorkSheetGlobals(m_sheets[m_SheetIndex], out idx, out m_currentRowRecord))
			{
				//read next sheet
				m_SheetIndex++;
				initializeSheetRead();
				return;
			};

			if (idx == null)
			{
				//no index, but should have the first row record
				m_noIndex = true;
			}
			else			
			{
				m_dbCellAddrs = idx.DbCellAddresses;
				m_dbCellAddrsIndex = 0;
				m_cellOffset = findFirstDataCellOffset((int)m_dbCellAddrs[m_dbCellAddrsIndex]);
				if (m_cellOffset < 0)
				{
					fail("Badly formed binary file. Has INDEX but no DBCELL");
					return;
				}
					
			}
			
		}


		private void fail(string message)
		{
			m_exceptionMessage = message;
			m_isValid = false;

			m_file.Dispose();
			m_isClosed = true;

			m_sheets = null;
			m_stream = null;
			m_globals = null;
			m_encoding = null;
			m_hdr = null;
		}

		private  object tryConvertOADateTime(double value, ushort XFormat)
		{
		    ushort format = 0;
            if (XFormat >= 0 && XFormat < m_globals.ExtendedFormats.Count)
            {
                var rec = m_globals.ExtendedFormats[XFormat];
                switch (rec.ID)
                {
                    case BIFFRECORDTYPE.XF_V2:
                        format = (ushort) (rec.ReadByte(2) & 0x3F);
                        break;
                    case BIFFRECORDTYPE.XF_V3:
                          if ((rec.ReadByte(3) & 4) == 0)
                            return value;
                        format = rec.ReadByte(1);
                        break;
                    case BIFFRECORDTYPE.XF_V4:
                        if ((rec.ReadByte(5) & 4) == 0)
                            return value;
                        format = rec.ReadByte(1);
                        break;
                    
                    default:
                        if ((rec.ReadByte(m_globals.Sheets[m_globals.Sheets.Count-1].IsV8 ? 9 : 7) & 4) == 0)
                            return value;
                        
                        format = rec.ReadUInt16(2);
                        break;
                }
            }
			else
            {
            	format = XFormat;
            }

           
            switch (format)
            {
                // numeric built in formats
                case 0: //"General";
                case 1: //"0";
                case 2: //"0.00";
                case 3: //"#,##0";
                case 4: //"#,##0.00";
                case 5: //"\"$\"#,##0_);(\"$\"#,##0)";
                case 6: //"\"$\"#,##0_);[Red](\"$\"#,##0)";
                case 7: //"\"$\"#,##0.00_);(\"$\"#,##0.00)";
                case 8: //"\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
                case 9: //"0%";
                case 10: //"0.00%";
                case 11: //"0.00E+00";
                case 12: //"# ?/?";
                case 13: //"# ??/??";
                case 0x30:// "##0.0E+0";
               
                case 0x25:// "_(#,##0_);(#,##0)";
                case 0x26:// "_(#,##0_);[Red](#,##0)";
                case 0x27:// "_(#,##0.00_);(#,##0.00)";
                case 40:// "_(#,##0.00_);[Red](#,##0.00)";
                case 0x29:// "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)";
                case 0x2a:// "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)";
                case 0x2b:// "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)";
                case 0x2c:// "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)";
                    return value;

                    // date formats
                case 14: //this.GetDefaultDateFormat();
                case 15: //"D-MM-YY";
                case 0x10: // "D-MMM";
                case 0x11: // "MMM-YY";
                case 0x12: // "h:mm AM/PM";
                case 0x13: // "h:mm:ss AM/PM";
                case 20: // "h:mm";
                case 0x15: // "h:mm:ss";
                case 0x16: // string.Format("{0} {1}", this.GetDefaultDateFormat(), this.GetDefaultTimeFormat());
                   
                case 0x2d: // "mm:ss";
                case 0x2e: // "[h]:mm:ss";
                case 0x2f: // "mm:ss.0";
                    return Helpers.ConvertFromOATime(value);
                 case 0x31:// "@";
                    return value.ToString();

                default:
                    XlsBiffFormatString fmtString;
                    if (m_globals.Formats.TryGetValue(format, out fmtString) )
                    {
                        var fmt = fmtString.Value;
                    	var formatReader = new FormatReader() {FormatString = fmt};
						if (formatReader.IsDateFormatString())
                            return Helpers.ConvertFromOATime(value); 

                    }
                    return value;



            }

		    
		}

	    private  object tryConvertOADateTime(object value, ushort XFormat)
	    {
	        double _dValue;

	        if (value == null)
	            return null;

			if (double.TryParse(value.ToString(), out _dValue))
				return tryConvertOADateTime(_dValue, XFormat);

			return value;

	    }

	    public bool isV8()
		{
			return m_version >= 0x600;
		}

		#endregion

		#region IExcelDataReader Members

		public void Initialize(Stream fileStream)
		{
			m_file = fileStream;

            readWorkBookGlobals();

            // set the sheet index to the index of the first sheet.. this is so that properties such as Name which use m_sheetIndex reflect the first sheet in the file without having to perform a read() operation
            m_SheetIndex = 0;
		}

		public void Reset() {
			m_SheetIndex = 0;
			m_IsFirstRead = true;
		}

		public string ExceptionMessage
		{
			get { return m_exceptionMessage; }
		}

		public string Name
		{
			get
			{
				if (null != m_sheets && m_sheets.Count > 0)
					return m_sheets[m_SheetIndex].Name;
				else
					return null;
			}
		}

        public string VisibleState
        {
            get
            {
                if (null != m_sheets && m_sheets.Count > 0)
                    return m_sheets[m_SheetIndex].VisibleState;
                else
                    return null;
            }
        }

		public bool IsValid
		{
			get { return m_isValid; }
		}

		public void Close()
		{
			m_file.Dispose();
			m_isClosed = true;
		}

		public int Depth
		{
			get { return m_depth; }
		}

		public int ResultsCount
		{
			get { return m_globals.Sheets.Count; }
		}

		public bool IsClosed
		{
			get { return m_isClosed; }
		}

		public bool NextResult()
		{
			if (!m_isValid) return false;
			if (m_SheetIndex >= (this.ResultsCount - 1)) return false;

			m_SheetIndex++;

			m_IsFirstRead = true;

			return true;
		}

		public bool Read()
		{
			if (!m_isValid) return false;

			if (m_IsFirstRead) {
				initializeSheetRead();
				if (!m_isValid) return false;

				if (IsFirstRowAsColumnNames) {
					if (moveToNextRecord()) {
						m_cellsNames = new string[m_cellsValues.Length];
						for (var i = 0; i < m_cellsValues.Length; i++) {
							var value = m_cellsValues[i]?.ToString();
							if (value != null && value.Length > 0) { 
								m_cellsNames[i] = value.ToString();
							}
						}
					} else {
						return false;
					}
				} else {
					m_cellsNames = null; // no columns
				}
			}

			return moveToNextRecord();
		}

		public int FieldCount
		{
			get { return m_maxCol; }
		}

		public bool GetBoolean(int i)
		{
			if (IsDBNull(i)) return false;

			return Boolean.Parse(m_cellsValues[i].ToString());
		}

		public DateTime GetDateTime(int i)
		{
			if (IsDBNull(i)) return DateTime.MinValue;

            // requested change: 3
			object val = m_cellsValues[i];

            if (val is DateTime)
            {
                // if the value is already a datetime.. return it without further conversion
                return (DateTime)val;
            }

            // otherwise proceed with conversion attempts
            string valString = val.ToString();
            double dVal;

			try
			{
                dVal = double.Parse(valString);
			}
			catch (FormatException)
			{
                return DateTime.Parse(valString);
			}

			return DateTimeHelper.FromOADate(dVal);
		}

		public decimal GetDecimal(int i)
		{
			if (IsDBNull(i)) return decimal.MinValue;

			return decimal.Parse(m_cellsValues[i].ToString());
		}

		public double GetDouble(int i)
		{
			if (IsDBNull(i)) return double.MinValue;

			return double.Parse(m_cellsValues[i].ToString());
		}

		public float GetFloat(int i)
		{
			if (IsDBNull(i)) return float.MinValue;

			return float.Parse(m_cellsValues[i].ToString());
		}

		public short GetInt16(int i)
		{
			if (IsDBNull(i)) return short.MinValue;

			return short.Parse(m_cellsValues[i].ToString());
		}

		public int GetInt32(int i)
		{
			if (IsDBNull(i)) return int.MinValue;

			return int.Parse(m_cellsValues[i].ToString());
		}

		public long GetInt64(int i)
		{
			if (IsDBNull(i)) return long.MinValue;

			return long.Parse(m_cellsValues[i].ToString());
		}

		public string GetString(int i)
		{
			if (IsDBNull(i)) return null;

			return m_cellsValues[i].ToString();
		}

		public object GetValue(int i)
		{
			return m_cellsValues[i];
		}

		public bool IsDBNull(int i)
		{
            return (null == m_cellsValues[i]);
		}

		public object this[int i]
		{
			get { return m_cellsValues[i]; }
		}

		#endregion

		#region  Not Supported IDataReader Members


        //public DataTable GetSchemaTable()
        //{
        //    throw new NotSupportedException();
        //}

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

        //public IDataReader GetData(int i)
        //{
        //    throw new NotSupportedException();
        //}

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
			if (m_cellsNames == null)
				return null;
			return m_cellsNames[i];
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

		#endregion

		#region IExcelDataReader Members

	    public bool IsFirstRowAsColumnNames { get; set; }

	    public bool ConvertOaDate
	    {
	        get { return convertOaDate; }
	        set { convertOaDate = value; }
	    }

	    public ReadOption ReadOption
		{
			get { return readOption; }
            set { readOption = value;  }
		}

	    public Encoding Encoding
	    {
	        get { return m_encoding; }
	    }

	    public Encoding DefaultEncoding
	    {
	        get { return Encoding.UTF8; }
	    }

	    #endregion

    }

    internal class EncodingHelper
    {

        public static Encoding GetEncoding(ushort codePage)
        {
            var encoding = (Encoding)null;
            switch (codePage)
            {
                case 037: encoding = Encoding.GetEncoding("IBM037"); break;
                case 437: encoding = Encoding.GetEncoding("IBM437"); break;
                case 500: encoding = Encoding.GetEncoding("IBM500"); break;
                case 708: encoding = Encoding.GetEncoding("ASMO-708"); break;
                case 709: encoding = Encoding.GetEncoding(""); break;
                case 710: encoding = Encoding.GetEncoding(""); break;
                case 720: encoding = Encoding.GetEncoding("DOS-720"); break;
                case 737: encoding = Encoding.GetEncoding("ibm737"); break;
                case 775: encoding = Encoding.GetEncoding("ibm775"); break;
                case 850: encoding = Encoding.GetEncoding("ibm850"); break;
                case 852: encoding = Encoding.GetEncoding("ibm852"); break;
                case 855: encoding = Encoding.GetEncoding("IBM855"); break;
                case 857: encoding = Encoding.GetEncoding("ibm857"); break;
                case 858: encoding = Encoding.GetEncoding("IBM00858"); break;
                case 860: encoding = Encoding.GetEncoding("IBM860"); break;
                case 861: encoding = Encoding.GetEncoding("ibm861"); break;
                case 862: encoding = Encoding.GetEncoding("DOS-862"); break;
                case 863: encoding = Encoding.GetEncoding("IBM863"); break;
                case 864: encoding = Encoding.GetEncoding("IBM864"); break;
                case 865: encoding = Encoding.GetEncoding("IBM865"); break;
                case 866: encoding = Encoding.GetEncoding("cp866"); break;
                case 869: encoding = Encoding.GetEncoding("ibm869"); break;
                case 870: encoding = Encoding.GetEncoding("IBM870"); break;
                case 874: encoding = Encoding.GetEncoding("windows-874"); break;
                case 875: encoding = Encoding.GetEncoding("cp875"); break;
                case 932: encoding = Encoding.GetEncoding("shift_jis"); break;
                case 936: encoding = Encoding.GetEncoding("gb2312"); break;
                case 949: encoding = Encoding.GetEncoding("ks_c_5601-1987"); break;
                case 950: encoding = Encoding.GetEncoding("big5"); break;
                case 1026: encoding = Encoding.GetEncoding("IBM1026"); break;
                case 1047: encoding = Encoding.GetEncoding("IBM01047"); break;
                case 1140: encoding = Encoding.GetEncoding("IBM01140"); break;
                case 1141: encoding = Encoding.GetEncoding("IBM01141"); break;
                case 1142: encoding = Encoding.GetEncoding("IBM01142"); break;
                case 1143: encoding = Encoding.GetEncoding("IBM01143"); break;
                case 1144: encoding = Encoding.GetEncoding("IBM01144"); break;
                case 1145: encoding = Encoding.GetEncoding("IBM01145"); break;
                case 1146: encoding = Encoding.GetEncoding("IBM01146"); break;
                case 1147: encoding = Encoding.GetEncoding("IBM01147"); break;
                case 1148: encoding = Encoding.GetEncoding("IBM01148"); break;
                case 1149: encoding = Encoding.GetEncoding("IBM01149"); break;
                case 1200: encoding = Encoding.GetEncoding("utf-16"); break;
                case 1201: encoding = Encoding.GetEncoding("unicodeFFFE"); break;
                case 1250: encoding = Encoding.GetEncoding("windows-1250"); break;
                case 1251: encoding = Encoding.GetEncoding("windows-1251"); break;
                case 1252: encoding = Encoding.GetEncoding("windows-1252"); break;
                case 1253: encoding = Encoding.GetEncoding("windows-1253"); break;
                case 1254: encoding = Encoding.GetEncoding("windows-1254"); break;
                case 1255: encoding = Encoding.GetEncoding("windows-1255"); break;
                case 1256: encoding = Encoding.GetEncoding("windows-1256"); break;
                case 1257: encoding = Encoding.GetEncoding("windows-1257"); break;
                case 1258: encoding = Encoding.GetEncoding("windows-1258"); break;
                case 1361: encoding = Encoding.GetEncoding("Johab"); break;
                case 10000: encoding = Encoding.GetEncoding("macintosh"); break;
                case 10001: encoding = Encoding.GetEncoding("x-mac-japanese"); break;
                case 10002: encoding = Encoding.GetEncoding("x-mac-chinesetrad"); break;
                case 10003: encoding = Encoding.GetEncoding("x-mac-korean"); break;
                case 10004: encoding = Encoding.GetEncoding("x-mac-arabic"); break;
                case 10005: encoding = Encoding.GetEncoding("x-mac-hebrew"); break;
                case 10006: encoding = Encoding.GetEncoding("x-mac-greek"); break;
                case 10007: encoding = Encoding.GetEncoding("x-mac-cyrillic"); break;
                case 10008: encoding = Encoding.GetEncoding("x-mac-chinesesimp"); break;
                case 10010: encoding = Encoding.GetEncoding("x-mac-romanian"); break;
                case 10017: encoding = Encoding.GetEncoding("x-mac-ukrainian"); break;
                case 10021: encoding = Encoding.GetEncoding("x-mac-thai"); break;
                case 10029: encoding = Encoding.GetEncoding("x-mac-ce"); break;
                case 10079: encoding = Encoding.GetEncoding("x-mac-icelandic"); break;
                case 10081: encoding = Encoding.GetEncoding("x-mac-turkish"); break;
                case 10082: encoding = Encoding.GetEncoding("x-mac-croatian"); break;
                case 12000: encoding = Encoding.GetEncoding("utf-32"); break;
                case 12001: encoding = Encoding.GetEncoding("utf-32BE"); break;
                case 20000: encoding = Encoding.GetEncoding("x-Chinese_CNS"); break;
                case 20001: encoding = Encoding.GetEncoding("x-cp20001"); break;
                case 20002: encoding = Encoding.GetEncoding("x_Chinese-Eten"); break;
                case 20003: encoding = Encoding.GetEncoding("x-cp20003"); break;
                case 20004: encoding = Encoding.GetEncoding("x-cp20004"); break;
                case 20005: encoding = Encoding.GetEncoding("x-cp20005"); break;
                case 20105: encoding = Encoding.GetEncoding("x-IA5"); break;
                case 20106: encoding = Encoding.GetEncoding("x-IA5-German"); break;
                case 20107: encoding = Encoding.GetEncoding("x-IA5-Swedish"); break;
                case 20108: encoding = Encoding.GetEncoding("x-IA5-Norwegian"); break;
                case 20127: encoding = Encoding.GetEncoding("us-ascii"); break;
                case 20261: encoding = Encoding.GetEncoding("x-cp20261"); break;
                case 20269: encoding = Encoding.GetEncoding("x-cp20269"); break;
                case 20273: encoding = Encoding.GetEncoding("IBM273"); break;
                case 20277: encoding = Encoding.GetEncoding("IBM277"); break;
                case 20278: encoding = Encoding.GetEncoding("IBM278"); break;
                case 20280: encoding = Encoding.GetEncoding("IBM280"); break;
                case 20284: encoding = Encoding.GetEncoding("IBM284"); break;
                case 20285: encoding = Encoding.GetEncoding("IBM285"); break;
                case 20290: encoding = Encoding.GetEncoding("IBM290"); break;
                case 20297: encoding = Encoding.GetEncoding("IBM297"); break;
                case 20420: encoding = Encoding.GetEncoding("IBM420"); break;
                case 20423: encoding = Encoding.GetEncoding("IBM423"); break;
                case 20424: encoding = Encoding.GetEncoding("IBM424"); break;
                case 20833: encoding = Encoding.GetEncoding("x-EBCDIC-KoreanExtended"); break;
                case 20838: encoding = Encoding.GetEncoding("IBM-Thai"); break;
                case 20866: encoding = Encoding.GetEncoding("koi8-r"); break;
                case 20871: encoding = Encoding.GetEncoding("IBM871"); break;
                case 20880: encoding = Encoding.GetEncoding("IBM880"); break;
                case 20905: encoding = Encoding.GetEncoding("IBM905"); break;
                case 20924: encoding = Encoding.GetEncoding("IBM00924"); break;
                case 20932: encoding = Encoding.GetEncoding("EUC-JP"); break;
                case 20936: encoding = Encoding.GetEncoding("x-cp20936"); break;
                case 20949: encoding = Encoding.GetEncoding("x-cp20949"); break;
                case 21025: encoding = Encoding.GetEncoding("cp1025"); break;
                case 21027: encoding = Encoding.GetEncoding(""); break;
                case 21866: encoding = Encoding.GetEncoding("koi8-u"); break;
                case 28591: encoding = Encoding.GetEncoding("iso-8859-1"); break;
                case 28592: encoding = Encoding.GetEncoding("iso-8859-2"); break;
                case 28593: encoding = Encoding.GetEncoding("iso-8859-3"); break;
                case 28594: encoding = Encoding.GetEncoding("iso-8859-4"); break;
                case 28595: encoding = Encoding.GetEncoding("iso-8859-5"); break;
                case 28596: encoding = Encoding.GetEncoding("iso-8859-6"); break;
                case 28597: encoding = Encoding.GetEncoding("iso-8859-7"); break;
                case 28598: encoding = Encoding.GetEncoding("iso-8859-8"); break;
                case 28599: encoding = Encoding.GetEncoding("iso-8859-9"); break;
                case 28603: encoding = Encoding.GetEncoding("iso-8859-13"); break;
                case 28605: encoding = Encoding.GetEncoding("iso-8859-15"); break;
                case 29001: encoding = Encoding.GetEncoding("x-Europa"); break;
                case 38598: encoding = Encoding.GetEncoding("iso-8859-8-i"); break;
                case 50220: encoding = Encoding.GetEncoding("iso-2022-jp"); break;
                case 50221: encoding = Encoding.GetEncoding("csISO2022JP"); break;
                case 50222: encoding = Encoding.GetEncoding("iso-2022-jp"); break;
                case 50225: encoding = Encoding.GetEncoding("iso-2022-kr"); break;
                case 50227: encoding = Encoding.GetEncoding("x-cp50227"); break;
                case 50229: encoding = Encoding.GetEncoding(""); break;
                case 50930: encoding = Encoding.GetEncoding(""); break;
                case 50931: encoding = Encoding.GetEncoding(""); break;
                case 50933: encoding = Encoding.GetEncoding(""); break;
                case 50935: encoding = Encoding.GetEncoding(""); break;
                case 50936: encoding = Encoding.GetEncoding(""); break;
                case 50937: encoding = Encoding.GetEncoding(""); break;
                case 50939: encoding = Encoding.GetEncoding(""); break;
                case 51932: encoding = Encoding.GetEncoding("euc-jp"); break;
                case 51936: encoding = Encoding.GetEncoding("EUC-CN"); break;
                case 51949: encoding = Encoding.GetEncoding("euc-kr"); break;
                case 51950: encoding = Encoding.GetEncoding(""); break;
                case 52936: encoding = Encoding.GetEncoding("hz-gb-2312"); break;
                case 54936: encoding = Encoding.GetEncoding("GB18030"); break;
                case 57002: encoding = Encoding.GetEncoding("x-iscii-de"); break;
                case 57003: encoding = Encoding.GetEncoding("x-iscii-be"); break;
                case 57004: encoding = Encoding.GetEncoding("x-iscii-ta"); break;
                case 57005: encoding = Encoding.GetEncoding("x-iscii-te"); break;
                case 57006: encoding = Encoding.GetEncoding("x-iscii-as"); break;
                case 57007: encoding = Encoding.GetEncoding("x-iscii-or"); break;
                case 57008: encoding = Encoding.GetEncoding("x-iscii-ka"); break;
                case 57009: encoding = Encoding.GetEncoding("x-iscii-ma"); break;
                case 57010: encoding = Encoding.GetEncoding("x-iscii-gu"); break;
                case 57011: encoding = Encoding.GetEncoding("x-iscii-pa"); break;
                case 65000: encoding = Encoding.GetEncoding("utf-7"); break;
                case 65001: encoding = Encoding.GetEncoding("utf-8"); break;
            }

            return encoding;
        }
    }

    /// <summary>
	/// Strict is as normal, Loose is more forgiving and will not cause an exception if a record size takes it beyond the end of the file. It will be trunacted in this case (SQl Reporting Services)
	/// </summary>
	public enum ReadOption
	{
		Strict,
		Loose
	}
}