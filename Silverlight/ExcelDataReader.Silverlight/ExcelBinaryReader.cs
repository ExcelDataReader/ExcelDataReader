namespace ExcelDataReader.Silverlight
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Text;
	using Core.BinaryFormat;
	using Data;

	/// <summary>
	/// ExcelDataReader Class
	/// </summary>
	public class ExcelBinaryReader : IExcelDataReader
	{
		#region Members
		private readonly Encoding _DefaultEncoding = Encoding.Unicode;

		private Stream _File;
		private XlsHeader _Header;
		private List<XlsWorksheet> _Sheets;
		private XlsBiffStream _BiffStream;
		private IWorkBook _WorkbookData;
		private XlsWorkbookGlobals _WorkbookGlobals;
		private ushort _Version;
		private bool _ConvertOaDate;
		private Encoding _Encoding;
		private object[] _CellsValues;
		private uint[] _DatabaseCellAddresses;
		private int _DatabaseCellAddressesIndex;
		private bool _CanRead;
		private int _SheetIndex;
		private int _CellOffset;
		private int _MaxRowIndex;
		private bool _CloseStreamOnFail;

		private bool _IsFirstRead;

		private const string Workbook = "Workbook";
		private const string Book = "Book";
		private const string Column = "Column";

		private bool _IsDisposed;

		#endregion

		public IWorkBookFactory WorkBookFactory { get; set; }

		internal ExcelBinaryReader()
		{
			_Encoding = _DefaultEncoding;
			_Version = 0x0600;
			IsValid = true;
			_SheetIndex = -1;
			_IsFirstRead = true;
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
			if (_IsDisposed) return;

			if (disposing)
			{
				//if (_WorkbookData != null) _WorkbookData.Dispose();

				if (_Sheets != null) _Sheets.Clear();
			}

			//_WorkbookData = null;
			_Sheets = null;
			_BiffStream = null;
			_WorkbookGlobals = null;
			_Encoding = null;
			_Header = null;

			_IsDisposed = true;
		}

		~ExcelBinaryReader()
		{
			Dispose(false);
		}

		#endregion

		#region Private methods

		private int FindFirstDataCellOffset(int startOffset)
		{
			var startCell = (XlsBiffDbCell)_BiffStream.ReadAt(startOffset);
			XlsBiffRow row;

			var offs = startCell.RowAddress;

			do
			{
				row = _BiffStream.ReadAt(offs) as XlsBiffRow;
				if (row == null) break;

				offs += row.Size;

			} while (true);

			return offs;
		}

		private void ReadWorkBookGlobals()
		{
			//Read Header
			try
			{
				_Header = XlsHeader.ReadHeader(_File);
			}
			catch (Exceptions.HeaderException ex)
			{
				Fail(ex.Message);
				return;
			}
			catch (FormatException ex)
			{
				Fail(ex.Message);
				return;
			}

			var dir = new XlsRootDirectory(_Header);
			var workbookEntry = dir.FindEntry(Workbook) ?? dir.FindEntry(Book);

			if (workbookEntry == null)
			{ Fail(Errors.ErrorStreamWorkbookNotFound); return; }

			if (workbookEntry.EntryType != STGTY.STGTY_STREAM)
			{ Fail(Errors.ErrorWorkbookIsNotStream); return; }

			_BiffStream = new XlsBiffStream(_Header, workbookEntry.StreamFirstSector);

			_WorkbookGlobals = new XlsWorkbookGlobals();

			_BiffStream.Seek(0, SeekOrigin.Begin);

			var rec = _BiffStream.Read();
			var bof = rec as XlsBiffBOF;

			if (bof == null || bof.Type != BIFFTYPE.WorkbookGlobals)
			{ Fail(Errors.ErrorWorkbookGlobalsInvalidData); return; }

			var sst = false;

			_Version = bof.Version;
			_Sheets = new List<XlsWorksheet>();

			while (null != (rec = _BiffStream.Read()))
			{
				switch (rec.ID)
				{
					case BIFFRECORDTYPE.INTERFACEHDR:
						_WorkbookGlobals.InterfaceHdr = (XlsBiffInterfaceHdr)rec;
						break;
					case BIFFRECORDTYPE.BOUNDSHEET:
						var sheet = (XlsBiffBoundSheet)rec;

						if (sheet.Type != XlsBiffBoundSheet.SheetType.Worksheet) break;

						sheet.IsV8 = IsV8();
						sheet.UseEncoding = _Encoding;

						_Sheets.Add(new XlsWorksheet(_WorkbookGlobals.Sheets.Count, sheet));
						_WorkbookGlobals.Sheets.Add(sheet);

						break;
					case BIFFRECORDTYPE.MMS:
						_WorkbookGlobals.MMS = rec;
						break;
					case BIFFRECORDTYPE.COUNTRY:
						_WorkbookGlobals.Country = rec;
						break;
					case BIFFRECORDTYPE.CODEPAGE:

						_WorkbookGlobals.CodePage = (XlsBiffSimpleValueRecord)rec;

						try
						{
							_Encoding = /*Encoding.GetEncoding(_WorkbookGlobals.CodePage.Value);*/ Encoding.Unicode;
						}
						catch (ArgumentException)
						{
							// Warning - Password protection
							// TODO: Attach to ILog
						}

						break;
					case BIFFRECORDTYPE.FONT:
					case BIFFRECORDTYPE.FONT_V34:
						_WorkbookGlobals.Fonts.Add(rec);
						break;
					case BIFFRECORDTYPE.FORMAT:
					case BIFFRECORDTYPE.FORMAT_V23:
						_WorkbookGlobals.Formats.Add(rec);
						break;
					case BIFFRECORDTYPE.XF:
					case BIFFRECORDTYPE.XF_V4:
					case BIFFRECORDTYPE.XF_V3:
					case BIFFRECORDTYPE.XF_V2:
						_WorkbookGlobals.ExtendedFormats.Add(rec);
						break;
					case BIFFRECORDTYPE.SST:
						_WorkbookGlobals.SST = (XlsBiffSST)rec;
						sst = true;
						break;
					case BIFFRECORDTYPE.CONTINUE:
						if (!sst) break;
						var contSst = (XlsBiffContinue)rec;
						_WorkbookGlobals.SST.Append(contSst);
						break;
					case BIFFRECORDTYPE.EXTSST:
						_WorkbookGlobals.ExtSST = rec;
						sst = false;
						break;
					case BIFFRECORDTYPE.PROTECT:
					case BIFFRECORDTYPE.PASSWORD:
					case BIFFRECORDTYPE.PROT4REVPASSWORD:
						//IsProtected
						break;
					case BIFFRECORDTYPE.EOF:
						if (_WorkbookGlobals.SST != null)
							_WorkbookGlobals.SST.ReadStrings();
						return;

					default:
						continue;
				}
			}
		}

		private bool ReadWorkSheetGlobals(XlsWorksheet sheet, out XlsBiffIndex idx)
		{
			_BiffStream.Seek((int)sheet.DataOffset, SeekOrigin.Begin);

			var bof = _BiffStream.Read() as XlsBiffBOF;

			idx = null;

			if (bof == null || bof.Type != BIFFTYPE.Worksheet) return false;

			idx = _BiffStream.Read() as XlsBiffIndex;

			if (null == idx) return false;


			idx.IsV8 = IsV8();

			XlsBiffRecord trec;
			XlsBiffDimensions dims = null;

			do
			{
				trec = _BiffStream.Read();
				if (trec.ID != BIFFRECORDTYPE.DIMENSIONS) continue;
				dims = (XlsBiffDimensions)trec;
				break;
			} while (trec.ID != BIFFRECORDTYPE.ROW);

			FieldCount = 256;

			if (dims != null)
			{
				dims.IsV8 = IsV8();
				FieldCount = dims.LastColumn - 1;
				sheet.Dimensions = dims;
			}

			_MaxRowIndex = (int)idx.LastExistingRow;

			if (idx.LastExistingRow <= idx.FirstExistingRow)
			{
				return false;
			}

			Depth = 0;

			return true;
		}

		private bool ReadWorkSheetRow()
		{
			_CellsValues = new object[FieldCount];

			while (_CellOffset < _BiffStream.Size)
			{
				var rec = _BiffStream.ReadAt(_CellOffset);
				_CellOffset += rec.Size;

				if ((rec is XlsBiffDbCell)) { break; }
				if (rec is XlsBiffEOF) { return false; }

				var cell = rec as XlsBiffBlankCell;

				if ((null == cell) || (cell.ColumnIndex >= FieldCount)) continue;
				if (cell.RowIndex != Depth) { _CellOffset -= rec.Size; break; }

				PushCellValue(cell);
			}

			Depth++;

			return Depth < _MaxRowIndex;
		}

		private IWorkSheet ReadWholeWorkSheet(XlsWorksheet sheet, IWorkBook workBook)
		{
			XlsBiffIndex idx;

			if (!ReadWorkSheetGlobals(sheet, out idx)) return null;

			var workSheet = workBook.CreateWorkSheet();
			workSheet.Name = sheet.Name;

			var triggerCreateColumns = true;

			_DatabaseCellAddresses = idx.DbCellAddresses;

			for (var index = 0; index < _DatabaseCellAddresses.Length; index++)
			{
				if (Depth == _MaxRowIndex) break;

				// init reading data
				_CellOffset = FindFirstDataCellOffset((int)_DatabaseCellAddresses[index]);

				//DataTable columns
				if (triggerCreateColumns)
				{
					if (IsFirstRowAsColumnNames && ReadWorkSheetRow())
					{
						for (var i = 0; i < FieldCount; i++)
						{
							if (_CellsValues[i] != null && _CellsValues[i].ToString().Length > 0)
							{
								var column = workSheet.CreateDataColumn();
								column.ColumnName = _CellsValues[i].ToString();
								workSheet.Columns.Add(column);
							}
							else
							{
								var column = workSheet.CreateDataColumn();
								column.ColumnName = String.Concat(Column, i);
								workSheet.Columns.Add(column);
							}
						}
					}
					else
					{
						for (var i = 0; i < FieldCount; i++)
						{
							workSheet.Columns.Add(workSheet.CreateDataColumn());
						}
					}

					triggerCreateColumns = false;

					//table.BeginLoadData();
				}

				while (ReadWorkSheetRow())
				{
					var dataRow = workSheet.CreateDataRow();
					dataRow.Values = _CellsValues;
					workSheet.Rows.Add(dataRow);
				}

				if (Depth > 0)
				{
					var dataRow = workSheet.CreateDataRow();
					dataRow.Values = _CellsValues;
					workSheet.Rows.Add(dataRow);
				}
			}

			//table.EndLoadData();
			return workSheet;
		}

		private void PushCellValue(XlsBiffBlankCell cell)
		{
			double dValue;

			switch (cell.ID)
			{
				case BIFFRECORDTYPE.INTEGER:
				case BIFFRECORDTYPE.INTEGER_OLD:
					_CellsValues[cell.ColumnIndex] = ((XlsBiffIntegerCell)cell).Value;
					break;
				case BIFFRECORDTYPE.NUMBER:
				case BIFFRECORDTYPE.NUMBER_OLD:

					dValue = ((XlsBiffNumberCell)cell).Value;

					_CellsValues[cell.ColumnIndex] = !_ConvertOaDate ?
																		dValue : TryConvertOaDateTime(dValue, cell.XFormat);

					break;
				case BIFFRECORDTYPE.LABEL:
				case BIFFRECORDTYPE.LABEL_OLD:
				case BIFFRECORDTYPE.RSTRING:
					_CellsValues[cell.ColumnIndex] = ((XlsBiffLabelCell)cell).Value;
					break;
				case BIFFRECORDTYPE.LABELSST:
					var tmp = _WorkbookGlobals.SST.GetString(((XlsBiffLabelSSTCell)cell).SSTIndex);
					_CellsValues[cell.ColumnIndex] = tmp;
					break;
				case BIFFRECORDTYPE.RK:

					dValue = ((XlsBiffRKCell)cell).Value;

					_CellsValues[cell.ColumnIndex] = !_ConvertOaDate
					                                 	? dValue
					                                 	: TryConvertOaDateTime(dValue, cell.XFormat);

					break;
				case BIFFRECORDTYPE.MULRK:

					var rkCell = (XlsBiffMulRKCell)cell;
					for (var j = cell.ColumnIndex; j <= rkCell.LastColumnIndex; j++)
					{
						_CellsValues[j] = !_ConvertOaDate
														? rkCell.GetValue(j)
														: TryConvertOaDateTime(rkCell.GetValue(j), rkCell.GetXF(j));
					}

					break;
				case BIFFRECORDTYPE.BLANK:
				case BIFFRECORDTYPE.BLANK_OLD:
				case BIFFRECORDTYPE.MULBLANK:
					// Skip blank cells

					break;
				case BIFFRECORDTYPE.FORMULA:
				case BIFFRECORDTYPE.FORMULA_OLD:

					var oValue = ((XlsBiffFormulaCell)cell).Value;

					if (null != oValue && oValue is FORMULAERROR)
					{
						oValue = null;
					}
					else
					{
						_CellsValues[cell.ColumnIndex] = !_ConvertOaDate ? oValue : TryConvertOaDateTime(oValue, (ushort)(cell.XFormat + 75));//date time offset
					}

					break;
				default:
					break;
			}
		}

		private bool MoveToNextRecord()
		{
			if (null == _DatabaseCellAddresses ||
				_DatabaseCellAddressesIndex == _DatabaseCellAddresses.Length ||
				Depth == _MaxRowIndex) return false;

			_CanRead = ReadWorkSheetRow();

			//read last row
			if (!_CanRead && Depth > 0) _CanRead = true;

			if (!_CanRead && _DatabaseCellAddressesIndex < (_DatabaseCellAddresses.Length - 1))
			{
				_DatabaseCellAddressesIndex++;
				_CellOffset = FindFirstDataCellOffset((int)_DatabaseCellAddresses[_DatabaseCellAddressesIndex]);

				_CanRead = ReadWorkSheetRow();
			}

			return _CanRead;
		}

		private void InitializeSheetRead()
		{
			if (_SheetIndex == ResultsCount) return;

			_DatabaseCellAddresses = null;

			_IsFirstRead = false;

			if (_SheetIndex == -1) _SheetIndex = 0;

			XlsBiffIndex idx;

			if (!ReadWorkSheetGlobals(_Sheets[_SheetIndex], out idx))
			{
				//read next sheet
				_SheetIndex++;
				InitializeSheetRead();
				return;
			}

			_DatabaseCellAddresses = idx.DbCellAddresses;
			_DatabaseCellAddressesIndex = 0;
			_CellOffset = FindFirstDataCellOffset((int)_DatabaseCellAddresses[_DatabaseCellAddressesIndex]);
		}

		private void Fail(string message)
		{
			ExceptionMessage = message;

			IsValid = false;
			if (_CloseStreamOnFail)
			{
				_File.Close();
			}

			IsClosed = true;

			//_WorkbookData = null;
			_Sheets = null;
			_BiffStream = null;
			_WorkbookGlobals = null;
			_Encoding = null;
			_Header = null;
		}

		private static object TryConvertOaDateTime(double value, ushort xFormat)
		{
			switch (xFormat)
			{
				//Time format
				case 63:
				case 68:
					var time = DateTime.FromOADate(value);

					return (time.Second == 0)
							? time.ToShortTimeString()
							: time.ToLongTimeString();

				//Date Format
				case 23: // region-specific format
				case 26:
				case 62:
				case 64:
				case 67:
				case 69:
				case 70:
				case 100: return DateTime.FromOADate(value).ToString(System.Globalization.CultureInfo.InvariantCulture);

				default:
					return value;
			}
		}

		private static object TryConvertOaDateTime(object value, ushort xFormat)
		{
			object r;

			try
			{
				var dValue = double.Parse(value.ToString());

				r = TryConvertOaDateTime(dValue, xFormat);
			}
			catch (FormatException)
			{
				r = value;
			}

			return r;
		}

		private bool IsV8()
		{
			return _Version >= 0x600;
		}

		#endregion

		#region IExcelDataReader Members

		public void Initialize(Stream fileStream)
		{
			Initialize(fileStream, true);
		}

		public void Initialize(Stream fileStream, bool closeOnFail)
		{
			_File = fileStream;
			_CloseStreamOnFail = closeOnFail;

			ReadWorkBookGlobals();
		}

		public IWorkBook AsWorkBook()
		{
			return AsWorkBook(false);
		}

		public IWorkBook AsWorkBook(bool convertOaDateTime)
		{
			if (!IsValid) return null;

			if (IsClosed) return _WorkbookData;

			_ConvertOaDate = convertOaDateTime;
			_WorkbookData = WorkBookFactory.CreateWorkBook();

			for (var index = 0; index < ResultsCount; index++)
			{
				var table = ReadWholeWorkSheet(_Sheets[index], _WorkbookData);

				if (null != table)
					_WorkbookData.WorkSheets.Add(table);
			}

			_File.Close();
			IsClosed = true;

			return _WorkbookData;
		}

		public string ExceptionMessage { get; private set; }

		public string Name
		{
			get
			{
				if (null != _Sheets && _Sheets.Count > 0)
					return _Sheets[_SheetIndex].Name;
				return null;
			}
		}

		public bool IsValid { get; private set; }

		public void Close()
		{
			_File.Close();
			IsClosed = true;
		}

		public int Depth { get; private set; }

		public int ResultsCount
		{
			get { return _WorkbookGlobals.Sheets.Count; }
		}

		public bool IsClosed { get; private set; }

		public bool NextResult()
		{
			if (_SheetIndex >= (ResultsCount - 1)) return false;

			_SheetIndex++;

			_IsFirstRead = true;

			return true;
		}

		public bool Read()
		{
			if (!IsValid) return false;

			if (_IsFirstRead) InitializeSheetRead();

			return MoveToNextRecord();
		}

		public int FieldCount { get; private set; }

		public bool GetBoolean(int i)
		{
			if (IsDBNull(i)) return false;

			return Boolean.Parse(_CellsValues[i].ToString());
		}

		public DateTime GetDateTime(int i)
		{
			if (IsDBNull(i)) return DateTime.MinValue;

			var val = _CellsValues[i].ToString();
			double dVal;

			try
			{
				dVal = double.Parse(val);
			}
			catch (FormatException)
			{
				return DateTime.Parse(val);
			}

			return DateTime.FromOADate(dVal);
		}

		public decimal GetDecimal(int i)
		{
			if (IsDBNull(i)) return decimal.MinValue;

			return decimal.Parse(_CellsValues[i].ToString());
		}

		public double GetDouble(int i)
		{
			if (IsDBNull(i)) return double.MinValue;

			return double.Parse(_CellsValues[i].ToString());
		}

		public float GetFloat(int i)
		{
			if (IsDBNull(i)) return float.MinValue;

			return float.Parse(_CellsValues[i].ToString());
		}

		public short GetInt16(int i)
		{
			if (IsDBNull(i)) return short.MinValue;

			return short.Parse(_CellsValues[i].ToString());
		}

		public int GetInt32(int i)
		{
			if (IsDBNull(i)) return int.MinValue;

			return int.Parse(_CellsValues[i].ToString());
		}

		public long GetInt64(int i)
		{
			if (IsDBNull(i)) return long.MinValue;

			return long.Parse(_CellsValues[i].ToString());
		}

		public string GetString(int i)
		{
			if (IsDBNull(i)) return null;

			return _CellsValues[i].ToString();
		}

		public object GetValue(int i)
		{
			return _CellsValues[i];
		}

		public bool IsDBNull(int i)
		{
			return (null == _CellsValues[i]) || (DBNull.Value == _CellsValues[i]);
		}

		public object this[int i]
		{
			get { return _CellsValues[i]; }
		}

		public bool IsFirstRowAsColumnNames { get; set; }

		#endregion
	}
}