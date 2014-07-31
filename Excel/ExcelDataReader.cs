using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using Excel.Core.BinaryFormat;

namespace Excel
{
	/// <summary>
	/// ExcelDataReader Class
	/// </summary>
	public class ExcelDataReader : IDisposable
	{
		#region Members

		private Stream m_file;
		private XlsHeader m_hdr;
		private List<XlsWorksheet> m_sheets;
		private XlsBiffStream m_stream;
		private DataSet m_workbookData;
		private XlsWorkbookGlobals m_globals;
		private ushort m_version;
		private bool m_PromoteToColumns;
		private bool m_ConvertOADate;
		private bool m_IsProtected;
		private Encoding m_encoding;

		private readonly Encoding m_Default_Encoding = Encoding.UTF8;

		private const string WORKBOOK = "Workbook";
		private const string BOOK = "Book";
		private const string COLUMN_DEFAULT_NAME = "Column";

		private bool disposed;

		#endregion

		#region Properties

		/// <summary>
		/// DataSet with workbook data, Tables represent Sheets
		/// </summary>
		public DataSet WorkbookData
		{
			get { return m_workbookData; }
		}

		/// <summary>
		/// Gets a value indicating whether the Xls file is protected.
		/// </summary>
		/// <value>
		/// 	<c>true</c> if this Xls file is proteted; otherwise, <c>false</c>.
		/// </value>
		public bool IsProtected
		{
			get { return m_IsProtected; }
		}

		#endregion

		#region Constructor and IDisposable Members

		/// <summary>
		/// Initializes a new instance of the <see cref="ExcelDataReader"/> class.
		/// </summary>
		/// <param name="fileStream">Xls Stream</param>
		public ExcelDataReader(Stream fileStream)
			: this(fileStream, false, true)
		{
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="ExcelDataReader"/> class.
		/// </summary>
		/// <param name="fileStream">Xls Stream</param>
		/// <param name="promoteToColumns">if is set to <c>true</c> the first row will be moved in to column names.</param>
		/// <param name="convertOADate">if is set to <c>true</c> the Date and Time values from the Xls Stream will be mapped as strings in the output DataSet.</param>
		public ExcelDataReader(Stream fileStream, bool promoteToColumns, bool convertOADate)
		{
			m_PromoteToColumns = promoteToColumns;
			m_encoding = m_Default_Encoding;
			m_version = 0x0600;
			m_ConvertOADate = convertOADate;
			m_sheets = new List<XlsWorksheet>();

			ParseXlsStream(fileStream);
		}

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
					m_workbookData.Dispose();

					m_sheets.Clear();
				}

				m_workbookData = null;
				m_sheets = null;
				m_stream = null;
				m_globals = null;
				m_encoding = null;
				m_hdr = null;

				disposed = true;
			}
		}

		~ExcelDataReader()
		{
			Dispose(false);
		}

		#endregion

		#region Private methods

		private void ParseXlsStream(Stream fileStream)
		{
			using (m_file = fileStream)
			{
				m_hdr = XlsHeader.ReadHeader(m_file);
				XlsRootDirectory dir = new XlsRootDirectory(m_hdr);
				XlsDirectoryEntry workbookEntry = dir.FindEntry(WORKBOOK) ?? dir.FindEntry(BOOK);

				if (workbookEntry == null)
					throw new FileNotFoundException(Errors.ErrorStreamWorkbookNotFound);
				if (workbookEntry.EntryType != STGTY.STGTY_STREAM)
					throw new FormatException(Errors.ErrorWorkbookIsNotStream);

				m_stream = new XlsBiffStream(m_hdr, workbookEntry.StreamFirstSector);

				ReadWorkbookGlobals();

				m_workbookData = new DataSet();

				for (int i = 0; i < m_sheets.Count; i++)
				{
					if (ReadWorksheet(m_sheets[i]))
						m_workbookData.Tables.Add(m_sheets[i].Data);
				}

				m_globals.SST = null;
				m_globals = null;
				m_sheets = null;
				m_stream = null;
				m_hdr = null;

				GC.Collect();
				GC.SuppressFinalize(this);
			}
		}

		private void ReadWorkbookGlobals()
		{
			m_globals = new XlsWorkbookGlobals();

			m_stream.Seek(0, SeekOrigin.Begin);

			XlsBiffRecord rec = m_stream.Read();
			XlsBiffBOF bof = rec as XlsBiffBOF;

			if (bof == null || bof.Type != BIFFTYPE.WorkbookGlobals)
				throw new ArgumentException(Errors.ErrorWorkbookGlobalsInvalidData);

			m_version = bof.Version;

			bool sst = false;

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

						sheet.IsV8 = IsV8();
						sheet.UseEncoding = m_encoding;

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

						try
						{
							m_encoding = Encoding.GetEncoding(m_globals.CodePage.Value);
						}
						catch
						{
							// Warning - Password protection
							// TODO: Attach to ILog
						}

						break;
					case BIFFRECORDTYPE.FONT:
					case BIFFRECORDTYPE.FONT_V34:
						m_globals.Fonts.Add(rec);
						break;
					case BIFFRECORDTYPE.FORMAT:
					case BIFFRECORDTYPE.FORMAT_V23:
						m_globals.Formats.Add(rec);
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
						m_IsProtected = true;
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

		private bool ReadWorksheet(XlsWorksheet sheet)
		{
			m_stream.Seek((int)sheet.DataOffset, SeekOrigin.Begin);

			XlsBiffBOF bof = m_stream.Read() as XlsBiffBOF;
			if (bof == null || bof.Type != BIFFTYPE.Worksheet)
				return false;

			XlsBiffIndex idx = m_stream.Read() as XlsBiffIndex;

			if (null == idx) return false;

			idx.IsV8 = IsV8();

			DataTable dt = new DataTable(sheet.Name);

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

			int maxCol = 256;

			if (dims != null)
			{
				dims.IsV8 = IsV8();
				maxCol = dims.LastColumn - 1;
				sheet.Dimensions = dims;
			}

			InitializeColumns(ref dt, maxCol);

			sheet.Data = dt;

			uint maxRow = idx.LastExistingRow;
			if (idx.LastExistingRow <= idx.FirstExistingRow)
			{
				return true;
			}

			dt.BeginLoadData();

			for (int i = 0; i < maxRow; i++)
			{
				dt.Rows.Add(dt.NewRow());
			}

			uint[] dbCellAddrs = idx.DbCellAddresses;

			for (int i = 0; i < dbCellAddrs.Length; i++)
			{
				XlsBiffDbCell dbCell = (XlsBiffDbCell)m_stream.ReadAt((int)dbCellAddrs[i]);
				XlsBiffRow row = null;
				int offs = dbCell.RowAddress;

				do
				{
					row = m_stream.ReadAt(offs) as XlsBiffRow;
					if (row == null) break;

					offs += row.Size;

				} while (null != row);

				while (true)
				{
					XlsBiffRecord rec = m_stream.ReadAt(offs);
					offs += rec.Size;
					if (rec is XlsBiffDbCell) break;
					if (rec is XlsBiffEOF) break;
					XlsBiffBlankCell cell = rec as XlsBiffBlankCell;

					if (cell == null) continue;
					if (cell.ColumnIndex >= maxCol) continue;
					if (cell.RowIndex > maxRow) continue;

					string _sValue;
					double _dValue;

					switch (cell.ID)
					{
						case BIFFRECORDTYPE.INTEGER:
						case BIFFRECORDTYPE.INTEGER_OLD:
							dt.Rows[cell.RowIndex][cell.ColumnIndex] = ((XlsBiffIntegerCell)cell).Value.ToString();
							break;
						case BIFFRECORDTYPE.NUMBER:
						case BIFFRECORDTYPE.NUMBER_OLD:

							_dValue = ((XlsBiffNumberCell)cell).Value;

							if ((_sValue = TryConvertOADate(_dValue, cell.XFormat)) != null)
							{
								dt.Rows[cell.RowIndex][cell.ColumnIndex] = _sValue;
							}
							else
							{
								dt.Rows[cell.RowIndex][cell.ColumnIndex] = _dValue;
							}

							break;
						case BIFFRECORDTYPE.LABEL:
						case BIFFRECORDTYPE.LABEL_OLD:
						case BIFFRECORDTYPE.RSTRING:
							dt.Rows[cell.RowIndex][cell.ColumnIndex] = ((XlsBiffLabelCell)cell).Value;
							break;
						case BIFFRECORDTYPE.LABELSST:
							string tmp = m_globals.SST.GetString(((XlsBiffLabelSSTCell)cell).SSTIndex);
							dt.Rows[cell.RowIndex][cell.ColumnIndex] = tmp;
							break;
						case BIFFRECORDTYPE.RK:

							_dValue = ((XlsBiffRKCell)cell).Value;

							if ((_sValue = TryConvertOADate(_dValue, cell.XFormat)) != null)
							{
								dt.Rows[cell.RowIndex][cell.ColumnIndex] = _sValue;
							}
							else
							{
								dt.Rows[cell.RowIndex][cell.ColumnIndex] = _dValue;
							}

							break;
						case BIFFRECORDTYPE.MULRK:

							XlsBiffMulRKCell _rkCell = (XlsBiffMulRKCell)cell;
							for (ushort j = cell.ColumnIndex; j <= _rkCell.LastColumnIndex; j++)
							{
								dt.Rows[cell.RowIndex][j] = _rkCell.GetValue(j);
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

							if (null != _oValue && _oValue is FORMULAERROR) _oValue = null;

							if (null != _oValue
								&& (_sValue = TryConvertOADate(_oValue, cell.XFormat)) != null)
							{
								dt.Rows[cell.RowIndex][cell.ColumnIndex] = _sValue;
							}
							else
							{
								dt.Rows[cell.RowIndex][cell.ColumnIndex] = _oValue;
							}

							break;
						default:
							break;
					}
				}
			}

			dt.EndLoadData();

			if (m_PromoteToColumns)
			{
				RemapColumnsNames(ref dt, dt.Rows[0].ItemArray);
				dt.Rows.RemoveAt(0);
				dt.AcceptChanges();
			}

			return true;
		}

		private string TryConvertOADate(double value, ushort XFormat)
		{
			if (!m_ConvertOADate) return null;

			switch (XFormat)
			{
				//Time format
				case 63:
				case 68:
					DateTime time = DateTime.FromOADate(value);

					return (time.Second == 0)
						? time.ToShortTimeString()
						: time.ToLongTimeString();

				//Date Format
				case 62:
				case 64:
				case 67:
				case 69:
				case 70: return DateTime.FromOADate(value).ToShortDateString();

				default:
					return null;
			}
		}

		private string TryConvertOADate(object value, ushort XFormat)
		{
			if (!m_ConvertOADate || null == value) return null;

			double _dValue;
			string _re;

			try
			{
				_dValue = double.Parse(value.ToString());

				_re = TryConvertOADate(_dValue, XFormat);
			}
			catch
			{
				_re = null;
			}


			return _re;

		}

		private static void InitializeColumns(ref DataTable dataTable, int columnsCount)
		{
			for (int i = 0; i < columnsCount; i++)
			{
				dataTable.Columns.Add(COLUMN_DEFAULT_NAME + (i + 1), typeof(string));
			}
		}

		private static void RemapColumnsNames(ref DataTable dataTable, object[] columnNames)
		{
			for (int index = 0; index < columnNames.Length; index++)
			{
				if (!string.IsNullOrEmpty(columnNames[index].ToString().Trim()))
				{
					dataTable.Columns[index].ColumnName = columnNames[index].ToString().Trim();
				}
			}
		}

		private bool IsV8()
		{
			return m_version >= 0x600;
		}

		#endregion

	}
}