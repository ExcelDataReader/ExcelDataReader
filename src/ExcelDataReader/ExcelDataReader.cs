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
                FallbackEncoding = configuration.FallbackEncoding
            };
        }

        ~ExcelDataReader()
        {
            Dispose(false);
        }

        public Encoding Encoding => Workbook?.Encoding;

        public string Name => _worksheetIterator?.Current?.Name;

        public string VisibleState => _worksheetIterator?.Current?.VisibleState;

        public int Depth { get; private set; }

        public int ResultsCount => Workbook?.ResultsCount ?? -1;

        public bool IsClosed { get; private set; }

        public int FieldCount => _worksheetIterator?.Current?.FieldCount ?? 0;

        public int RecordsAffected => throw new NotSupportedException();

        protected ExcelReaderConfiguration Configuration { get; }

        protected TWorkbook Workbook { get; set; }

        private object[] CellsValues => _rowIterator?.Current;

        public object this[int i] => CellsValues[i];

        public object this[string name] => throw new NotSupportedException();

        public bool GetBoolean(int i)
        {
            return !IsDBNull(i) && bool.Parse(CellsValues[i].ToString());
        }

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

        public DateTime GetDateTime(int i)
        {
            if (IsDBNull(i))
                return DateTime.MinValue;

            // requested change: 3
            object val = CellsValues[i];

            if (val is DateTime)
            {
                // if the value is already a datetime.. return it without further conversion
                return (DateTime)val;
            }

            return DateTime.MinValue;
        }

        public decimal GetDecimal(int i)
        {
            return IsDBNull(i) ? decimal.MinValue : decimal.Parse(CellsValues[i].ToString());
        }

        public double GetDouble(int i)
        {
            return IsDBNull(i) ? double.MinValue : double.Parse(CellsValues[i].ToString());
        }

        public Type GetFieldType(int i)
        {
            return CellsValues[i]?.GetType();
        }

        public float GetFloat(int i)
        {
            return IsDBNull(i) ? float.MinValue : float.Parse(CellsValues[i].ToString());
        }

        public Guid GetGuid(int i)
        {
            throw new NotSupportedException();
        }

        public short GetInt16(int i)
        {
            return IsDBNull(i) ? short.MinValue : short.Parse(CellsValues[i].ToString());
        }

        public int GetInt32(int i)
        {
            return IsDBNull(i) ? int.MinValue : int.Parse(CellsValues[i].ToString());
        }

        public long GetInt64(int i)
        {
            return IsDBNull(i) ? long.MinValue : long.Parse(CellsValues[i].ToString());
        }

        public string GetName(int i)
        {
            throw new NotSupportedException();
        }

        public int GetOrdinal(string name)
        {
            throw new NotSupportedException();
        }

        /// <inheritdoc />
        public DataTable GetSchemaTable()
        {
            throw new NotSupportedException();
        }

        public string GetString(int i)
        {
            return IsDBNull(i) ? null : CellsValues[i].ToString();
        }

        public object GetValue(int i)
        {
            return CellsValues[i];
        }

        public int GetValues(object[] values)
        {
            throw new NotSupportedException();
        }

        public bool IsDBNull(int i)
        {
            return CellsValues[i] == null;
        }

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
