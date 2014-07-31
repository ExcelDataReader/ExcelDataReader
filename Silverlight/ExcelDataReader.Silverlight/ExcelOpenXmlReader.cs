namespace ExcelDataReader.Silverlight
{
	using System;
	using System.Collections.Generic;
	using System.Globalization;
	using System.IO;
	using System.Xml;
	using Core;
	using Core.OpenXmlFormat;
	using Data;

	public class ExcelOpenXmlReader : IExcelDataReader
	{
		#region Members

		private const string Column = "Column";

		private readonly List<int> _defaultDateTimeStyles;
		private object[] _cellsValues;
		private int _emptyRowCount;

		private bool _isDisposed;
		private bool _isFirstRead;
		private int _resultIndex;
		private object[] _savedCellsValues;
		private Stream _sheetStream;
		private XlsxWorkbook _workbook;
		private XmlReader _xmlReader;
		private ZipWorker _zipWorker;

		#endregion

		internal ExcelOpenXmlReader()
		{
			IsValid = true;
			_isFirstRead = true;

			_defaultDateTimeStyles = new List<int>(new[]
			                                       	{
			                                       		14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
			                                       	});
		}

		public int Depth { get; private set; }

		public bool IsClosed { get; private set; }

		public IWorkBookFactory WorkBookFactory { get; set; }

		public int FieldCount
		{
			get { return (_resultIndex >= 0 && _resultIndex < ResultsCount) ? _workbook.Sheets[_resultIndex].ColumnsCount : -1; }
		}

		public object this[int i]
		{
			get { return _cellsValues[i]; }
		}

		public void Initialize(Stream fileStream)
		{
			Initialize(fileStream, true);
		}

		public void Initialize(Stream fileStream, bool closeOnFail)
		{
			_zipWorker = new ZipWorker();
			_zipWorker.Extract(fileStream);

			if (_zipWorker.IsValid)
			{
				ReadGlobals();
				fileStream.Close();
			}
			else
			{
				IsValid = false;
				ExceptionMessage = _zipWorker.ExceptionMessage;

				Close();
				if (closeOnFail) fileStream.Close();
			}
		}

		public IWorkBook AsWorkBook()
		{
			return AsWorkBook(true);
		}

		public IWorkBook AsWorkBook(bool convertOaDateTime)
		{
			if (!IsValid) return null;

			var workBook = WorkBookFactory.CreateWorkBook();

			for (_resultIndex = 0; _resultIndex < _workbook.Sheets.Count; _resultIndex++)
			{
				var workSheet = workBook.CreateWorkSheet();
				workSheet.Name =_workbook.Sheets[_resultIndex].Name;

				ReadSheetGlobals(_workbook.Sheets[_resultIndex]);

				if (_workbook.Sheets[_resultIndex].Dimension == null) continue;

				Depth = 0;
				_emptyRowCount = 0;

				//DataTable columns
				if (!IsFirstRowAsColumnNames)
				{
					for (var i = 0; i < _workbook.Sheets[_resultIndex].ColumnsCount; i++)
					{
						workSheet.Columns.Add(workSheet.CreateDataColumn());
					}
				}
				else if (ReadSheetRow(_workbook.Sheets[_resultIndex]))
				{
					for (var index = 0; index < _cellsValues.Length; index++)
					{
						if (_cellsValues[index] != null && _cellsValues[index].ToString().Length > 0)
						{
							var dataColumn = workSheet.CreateDataColumn();
							dataColumn.ColumnName = _cellsValues[index].ToString();
							workSheet.Columns.Add(dataColumn);
						}
						else
						{
							var dataColumn = workSheet.CreateDataColumn();
							dataColumn.ColumnName = String.Concat(Column, index);
							workSheet.Columns.Add(dataColumn);
						}
					}
				}
				else continue;

				while (ReadSheetRow(_workbook.Sheets[_resultIndex]))
				{
					var row = workSheet.CreateDataRow();
					row.Values = _cellsValues;
					workSheet.Rows.Add(row);
				}

				if (workSheet.Rows.Count > 0)
					workBook.WorkSheets.Add(workSheet);
			}

			return workBook;
		}

		public bool IsFirstRowAsColumnNames { get; set; }

		public bool IsValid { get; private set; }

		public string ExceptionMessage { get; private set; }

		public string Name
		{
			get { return (_resultIndex >= 0 && _resultIndex < ResultsCount) ? _workbook.Sheets[_resultIndex].Name : null; }
		}

		public int ResultsCount
		{
			get { return _workbook == null ? -1 : _workbook.Sheets.Count; }
		}

		private void ReadGlobals()
		{
			_workbook = new XlsxWorkbook(
				_zipWorker.GetWorkbookByteArray(),
                _zipWorker.GetWorkbookRelsByteArray(),
                _zipWorker.GetSharedStringsByteArray(),
                _zipWorker.GetStylesByteArray());

			CheckDateTimeNumFmts(_workbook.Styles.NumFmts);
		}

		private void CheckDateTimeNumFmts(ICollection<XlsxNumFmt> list)
		{
			if (list.Count == 0) return;

			foreach (var numFmt in list)
			{
				if (string.IsNullOrEmpty(numFmt.FormatCode)) continue;
				var formatCode = numFmt.FormatCode;

				int index;
				while ((index = formatCode.IndexOf('"')) > 0)
				{
					var endPosition = formatCode.IndexOf('"', index + 1);

					if (endPosition > 0) formatCode = formatCode.Remove(index, endPosition - index + 1);
				}

				var firstSection = formatCode.Split(';')[0];
				var openBracketIndex = firstSection.IndexOf('[');
				var closeBracketIndex = firstSection.IndexOf(']');

				if (openBracketIndex >= 0 && closeBracketIndex >=0)
				{
					closeBracketIndex += 1;
					firstSection = firstSection.Substring(0, openBracketIndex) +
					               ((firstSection.Length > closeBracketIndex)
					                	? firstSection.Substring(closeBracketIndex)
					                	: String.Empty);
				}

				var dateChars = new char[] {'y', 'm', 'd', 's', 'h'};
				if (firstSection.IndexOfAny(dateChars) >= 0)
					_defaultDateTimeStyles.Add(numFmt.Id);
			}
		}

		private void ReadSheetGlobals(XlsxWorksheet sheet)
		{
			//_SheetStream = new MemoryStream(_ZipWorker.GetWorksheetByteArray(sheet.Id));
            _sheetStream = new MemoryStream(_zipWorker.GetWorksheetByteArray(sheet.Path));

			if (null == _sheetStream) return;

			_xmlReader = XmlReader.Create(_sheetStream);

			while (_xmlReader.Read())
			{
				if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.Name == XlsxWorksheet.N_dimension)
				{
					string dimValue = _xmlReader.GetAttribute(XlsxWorksheet.A_ref);

					if (dimValue.IndexOf(':') > 0)
						sheet.Dimension = new XlsxDimension(dimValue);
					else
					{
						_xmlReader.Close();
						_sheetStream.Close();
					}

					break;
				}
			}
		}

		private bool ReadSheetRow(XlsxWorksheet sheet)
		{
            if (null == _xmlReader) return false;

            if (_emptyRowCount != 0)
            {
                _cellsValues = new object[sheet.ColumnsCount];
                _emptyRowCount--;
                Depth++;

                return true;
            }

            if (_savedCellsValues != null)
            {
                _cellsValues = _savedCellsValues;
                _savedCellsValues = null;
                Depth++;

                return true;
            }

            if ((_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.Name == XlsxWorksheet.N_row) ||
                _xmlReader.ReadToFollowing(XlsxWorksheet.N_row))
            {
                _cellsValues = new object[sheet.ColumnsCount];

                int rowIndex = int.Parse(_xmlReader.GetAttribute(XlsxWorksheet.A_r));
                if (rowIndex != (Depth + 1))
                {
                    _emptyRowCount = rowIndex - Depth - 1;
                }
                bool hasValue = false;
                string a_s = String.Empty;
                string a_t = String.Empty;
                string a_r = String.Empty;
                int col = 0;
                int row = 0;

                while (_xmlReader.Read())
                {
                    if (_xmlReader.Depth == 2) break;

                    if (_xmlReader.NodeType == XmlNodeType.Element)
                    {
                        hasValue = false;

                        if (_xmlReader.Name == XlsxWorksheet.N_c)
                        {
                            a_s = _xmlReader.GetAttribute(XlsxWorksheet.A_s);
                            a_t = _xmlReader.GetAttribute(XlsxWorksheet.A_t);
                            a_r = _xmlReader.GetAttribute(XlsxWorksheet.A_r);
                            XlsxDimension.XlsxDim(a_r, out col, out row);
                        }
                        else if (_xmlReader.Name == XlsxWorksheet.N_v)
                        {
                            hasValue = true;
                        }
                    }

                    if (_xmlReader.NodeType == XmlNodeType.Text && hasValue)
                    {
                        object o = _xmlReader.Value;

                        if (null != a_t && a_t == XlsxWorksheet.A_s)
                        {
                            o = _workbook.SST[Convert.ToInt32(o)];
                        }
                        else if (null != a_s)
                        {
                            XlsxXf xf = _workbook.Styles.CellXfs[int.Parse(a_s)];

                            if (xf.ApplyNumberFormat && IsDateTimeStyle(xf.NumFmtId) && o != null && o.ToString() != string.Empty)
                            {
                                o = DateTime.FromOADate(Convert.ToDouble(o, CultureInfo.InvariantCulture));
                            }
                        }

                        if (col - 1 < _cellsValues.Length)
                            _cellsValues[col - 1] = o;
                    }
                }

                if (_emptyRowCount > 0)
                {
                    _savedCellsValues = _cellsValues;
                    return ReadSheetRow(sheet);
                }
                else
                    Depth++;

                return true;
            }

            _xmlReader.Close();
            if (_sheetStream != null) _sheetStream.Close();

            return false;
		}

		private bool InitializeSheetRead()
		{
			if (ResultsCount <= 0) return false;

			ReadSheetGlobals(_workbook.Sheets[_resultIndex]);

			if (_workbook.Sheets[_resultIndex].Dimension == null) return false;

			_isFirstRead = false;

			Depth = 0;
			_emptyRowCount = 0;

			return true;
		}

		private bool IsDateTimeStyle(int styleId)
		{
			return _defaultDateTimeStyles.Contains(styleId);
		}

		public void Close()
		{
			IsClosed = true;

			if (_xmlReader != null) _xmlReader.Close();

			if (_sheetStream != null) _sheetStream.Close();

			if (_zipWorker != null) _zipWorker.Dispose();
		}

		public bool NextResult()
		{
			if (_resultIndex >= (ResultsCount - 1)) return false;

			_resultIndex++;

			_isFirstRead = true;

			return true;
		}

		public bool Read()
		{
			if (!IsValid) return false;

			if (_isFirstRead && !InitializeSheetRead())
				return false;

			return ReadSheetRow(_workbook.Sheets[_resultIndex]);
		}

		public bool GetBoolean(int i)
		{
			if (IsDBNull(i)) return false;

			return Boolean.Parse(_cellsValues[i].ToString());
		}

		public DateTime GetDateTime(int i)
		{
			if (IsDBNull(i)) return DateTime.MinValue;

			try
			{
				return (DateTime) _cellsValues[i];
			}
			catch (InvalidCastException)
			{
				return DateTime.MinValue;
			}
		}

		public decimal GetDecimal(int i)
		{
			if (IsDBNull(i)) return decimal.MinValue;

			return decimal.Parse(_cellsValues[i].ToString());
		}

		public double GetDouble(int i)
		{
			if (IsDBNull(i)) return double.MinValue;

			return double.Parse(_cellsValues[i].ToString());
		}

		public float GetFloat(int i)
		{
			if (IsDBNull(i)) return float.MinValue;

			return float.Parse(_cellsValues[i].ToString());
		}

		public short GetInt16(int i)
		{
			if (IsDBNull(i)) return short.MinValue;

			return short.Parse(_cellsValues[i].ToString());
		}

		public int GetInt32(int i)
		{
			if (IsDBNull(i)) return int.MinValue;

			return int.Parse(_cellsValues[i].ToString());
		}

		public long GetInt64(int i)
		{
			if (IsDBNull(i)) return long.MinValue;

			return long.Parse(_cellsValues[i].ToString());
		}

		public string GetString(int i)
		{
			if (IsDBNull(i)) return null;

			return _cellsValues[i].ToString();
		}

		public object GetValue(int i)
		{
			return _cellsValues[i];
		}

		public bool IsDBNull(int i)
		{
			return (null == _cellsValues[i]) || (DBNull.Value == _cellsValues[i]);
		}

		public void Dispose()
		{
			Dispose(true);

			GC.SuppressFinalize(this);
		}

		private void Dispose(bool disposing)
		{
			// Check to see if Dispose has already been called.
			if (!_isDisposed)
			{
				if (disposing)
				{
					if (_zipWorker != null) _zipWorker.Dispose();
					if (_xmlReader != null) _xmlReader.Close();
					if (_sheetStream != null) _sheetStream.Close();
				}

				_zipWorker = null;
				_xmlReader = null;
				_sheetStream = null;

				_workbook = null;
				_cellsValues = null;
				_savedCellsValues = null;

				_isDisposed = true;
			}
		}

		~ExcelOpenXmlReader()
		{
			Dispose(false);
		}
	}
}