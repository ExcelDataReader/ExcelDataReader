using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using ExcelDataReader.Core;
using ExcelDataReader.Misc;

namespace ExcelDataReader
{
    /// <summary>
    /// A generic implementation of the IExcelDataReader interface using IWorkbook/IWorksheet to enumerate data.
    /// </summary>
    /// <typeparam name="TWorkbook">A type implementing IWorkbook</typeparam>
    /// <typeparam name="TWorksheet">A type implementing IWorksheet</typeparam>
    internal abstract class ExcelDataReader<TWorkbook, TWorksheet> : IExcelDataReader
        where TWorkbook : IWorkbook<TWorksheet>
        where TWorksheet : IWorksheet
    {
        private IEnumerator<TWorksheet> _worksheetIterator;
        private IEnumerator<object[]> _rowIterator;

        protected ExcelDataReader(ExcelReaderConfiguration configuration)
        {
            if (configuration == null)
            {
                // Use the defaults
                configuration = new ExcelReaderConfiguration();
            }

            // Copy the configuration to prevent external changes
            Configuration = new ExcelReaderConfiguration()
            {
                FallbackEncoding = configuration.FallbackEncoding,
                Password = configuration.Password
            };
        }

        ~ExcelDataReader()
        {
            Dispose(false);
        }

        public string Name => _worksheetIterator?.Current?.Name;

        public string CodeName => _worksheetIterator?.Current?.CodeName;

        public string VisibleState => _worksheetIterator?.Current?.VisibleState;

        public HeaderFooter HeaderFooter => _worksheetIterator?.Current?.HeaderFooter;
        
        public int Depth { get; private set; }

        public int ResultsCount => Workbook?.ResultsCount ?? -1;

        public bool IsClosed { get; private set; }

        public int FieldCount => _worksheetIterator?.Current?.FieldCount ?? 0;

        public int RecordsAffected => throw new NotSupportedException();

        protected ExcelReaderConfiguration Configuration { get; }

        protected TWorkbook Workbook { get; set; }

        private object[] CellsValues
        {
            get
            {
                if (_rowIterator == null || _rowIterator.Current == null)
                    throw new InvalidOperationException("No data exists for the row/column.");
                return _rowIterator?.Current;
            }
        }

        public object this[int i] => CellsValues[i];

        public object this[string name] => throw new NotSupportedException();

        public bool GetBoolean(int i) => (bool)CellsValues[i];

        public byte GetByte(int i) => (byte)CellsValues[i];

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
            => throw new NotSupportedException();

        public char GetChar(int i) => (char)CellsValues[i];

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
             => throw new NotSupportedException();

        public IDataReader GetData(int i) => throw new NotSupportedException();

        public string GetDataTypeName(int i) => throw new NotSupportedException();

        public DateTime GetDateTime(int i) => (DateTime)CellsValues[i];

        public decimal GetDecimal(int i) => (decimal)CellsValues[i];

        public double GetDouble(int i) => (double)CellsValues[i];

        public Type GetFieldType(int i) => CellsValues[i]?.GetType();

        public float GetFloat(int i) => (float)CellsValues[i];

        public Guid GetGuid(int i) => (Guid)CellsValues[i];

        public short GetInt16(int i) => (short)CellsValues[i];

        public int GetInt32(int i) => (int)CellsValues[i];

        public long GetInt64(int i) => (long)CellsValues[i];

        public string GetName(int i) => throw new NotSupportedException();

        public int GetOrdinal(string name) => throw new NotSupportedException();

        /// <inheritdoc />
        public DataTable GetSchemaTable() => throw new NotSupportedException();

        public string GetString(int i) => (string)CellsValues[i];

        public object GetValue(int i) => CellsValues[i];

        public int GetValues(object[] values) => throw new NotSupportedException();

        public bool IsDBNull(int i) => CellsValues[i] == null;

        /// <inheritdoc />
        public void Reset()
        {
            _worksheetIterator?.Dispose();
            _rowIterator?.Dispose();

            _worksheetIterator = null;
            _rowIterator = null;

            ResetSheetData();

            if (Workbook != null)
            {
                _worksheetIterator = Workbook.ReadWorksheets().GetEnumerator();
                if (!_worksheetIterator.MoveNext())
                {
                    _worksheetIterator.Dispose();
                    _worksheetIterator = null;
                    return;
                }

                _rowIterator = _worksheetIterator.Current.ReadRows().GetEnumerator();
            }
        }

        public virtual void Close()
        {
            if (IsClosed)
                return;

            _worksheetIterator?.Dispose();
            _rowIterator?.Dispose();

            _worksheetIterator = null;
            _rowIterator = null;
            IsClosed = true;
        }

        public bool NextResult()
        {
            if (_worksheetIterator == null)
            {
                return false;
            }

            ResetSheetData();

            _rowIterator?.Dispose();
            _rowIterator = null;

            if (!_worksheetIterator.MoveNext())
            {
                _worksheetIterator.Dispose();
                _worksheetIterator = null;
                return false;
            }

            _rowIterator = _worksheetIterator.Current.ReadRows().GetEnumerator();
            return true;
        }

        public bool Read()
        {
            if (_worksheetIterator == null || _rowIterator == null)
            {
                return false;
            }

            if (!_rowIterator.MoveNext())
            {
                _rowIterator.Dispose();
                _rowIterator = null;
                return false;
            }

            Depth++;
            return true;
        }
    
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

        private void ResetSheetData()
        {
            Depth = -1;
        }
    }
}
