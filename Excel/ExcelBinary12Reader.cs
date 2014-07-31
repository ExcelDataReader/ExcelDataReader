using System;
using System.Collections.Generic;
using System.Text;
using Excel.Core.Binary12Format;
using System.IO;
using Excel.Core;
using System.Data;

namespace Excel
{
	public class ExcelBinary12Reader : IExcelDataReader
	{
		#region Members

		private XlsbWorkbook _workbook;
		private bool _isValid;
		private bool _isClosed;
		private bool _isFirstRead;
		private string _exceptionMessage;
		private int _depth;
		private int _resultIndex;
		private int _emptyRowCount;
		private ZipWorker _zipWorker;

		private Stream _sheetStream;
		private object[] _cellsValues;
		private object[] _savedCellsValues;

		private bool disposed;

		#endregion

		internal ExcelBinary12Reader()
		{
			_isValid = true;
			_isFirstRead = true;
		}

		private void ReadGlobals()
		{
			_workbook = new XlsbWorkbook(
				_zipWorker.GetWorkbookStream(),
				_zipWorker.GetSharedStringsStream(),
				_zipWorker.GetStylesStream());
		}

		#region IExcelDataReader Members

		public void Initialize(System.IO.Stream fileStream)
		{
			_zipWorker = new ZipWorker(true);
			_zipWorker.Extract(fileStream);

			if (!_zipWorker.IsValid)
			{
				_isValid = false;
				_exceptionMessage = _zipWorker.ExceptionMessage;

				Close();

				return;
			}

			ReadGlobals();
		}


		public System.Data.DataSet AsDataSet()
		{
			return AsDataSet(false);
		}

		public System.Data.DataSet AsDataSet(bool convertOADateTime)
		{
			throw new Exception("The method or operation is not implemented.");
		}

		public bool IsValid
		{
			get { return _isValid; }
		}

		public string ExceptionMessage
		{
			get { return _exceptionMessage; }
		}

		public string Name
		{
			get
			{
				return (_resultIndex >= 0 && _resultIndex < ResultsCount) ? _workbook.Sheets[_resultIndex].Name : null;
			}
		}

		public int ResultsCount
		{
			get { return _workbook == null ? -1 : _workbook.Sheets.Count; }
		}

		#endregion

		#region IDataReader Members

		public void Close()
		{
			_isClosed = true;

			if (_sheetStream != null) _sheetStream.Close();

			if (_zipWorker != null) _zipWorker.Dispose();
		}

		public int Depth
		{
			get { return _depth; }
		}

		public bool IsClosed
		{
			get { return _isClosed; }
		}

		public bool NextResult()
		{
			throw new Exception("The method or operation is not implemented.");
		}

		public bool Read()
		{
			throw new Exception("The method or operation is not implemented.");
		}

		#endregion

		#region IDataRecord Members

		public int FieldCount
		{
			get { return (_resultIndex >= 0 && _resultIndex < ResultsCount) ? _workbook.Sheets[_resultIndex].ColumnsCount : -1; }

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
				return (DateTime)_cellsValues[i];
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

		public object this[int i]
		{
			get { return _cellsValues[i]; }
		}

		#endregion

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
					if (_zipWorker != null) _zipWorker.Dispose();
					if (_sheetStream != null) _sheetStream.Close();
				}

				_zipWorker = null;
				_sheetStream = null;

				_workbook = null;
				_cellsValues = null;
				_savedCellsValues = null;

				disposed = true;
			}
		}

		~ExcelBinary12Reader()
		{
			Dispose(false);
		}

		#endregion

		#region  Not Supported IDataReader Members

		public DataTable GetSchemaTable()
		{
			throw new NotSupportedException();
		}

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
			throw new NotSupportedException();
		}

		public Guid GetGuid(int i)
		{
			throw new NotSupportedException();
		}

		public string GetName(int i)
		{
			throw new NotSupportedException();
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


		public bool IsFirstRowAsColumnNames
		{
			get
			{
				throw new Exception("The method or operation is not implemented.");
			}
			set
			{
				throw new Exception("The method or operation is not implemented.");
			}
		}

		#endregion
	}
}
