using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using ExcelDataReader.Core.BinaryFormat;
using ExcelDataReader.Exceptions;
using ExcelDataReader.Misc;

namespace ExcelDataReader
{
    /// <summary>
    /// Strict is as normal, Loose is more forgiving and will not cause an exception if a record size takes it beyond the end of the file. It will be trunacted in this case (SQl Reporting Services)
    /// </summary>
    public enum ReadOption
    {
        Strict,
        Loose
    }

    /// <summary>
    /// ExcelDataReader Class
    /// </summary>
    public partial class ExcelBinaryReader : IExcelDataReader
    {
        private const string Workbook = "Workbook";
        private const string Book = "Book";
        private Stream _file;
        private XlsDocument _document;
        private XlsWorkbook _workbook;
        private bool _isFirstRead;
        private string[] _cellsNames;
        private IEnumerator<XlsWorksheet> _worksheetIterator = null;
        private IEnumerator<object[]> _rowIterator = null;

        public ExcelBinaryReader(Stream stream)
            : this(stream, true, ReadOption.Strict)
        {
        }

        public ExcelBinaryReader(Stream stream, ReadOption readOption)
            : this(stream, true, readOption)
        {
        }

        public ExcelBinaryReader(Stream stream, bool convertOADate, ReadOption readOption)
        {
            _file = stream;
            ReadOption = readOption;

            _document = new XlsDocument(stream);
            _workbook = ReadWorkbook(convertOADate);

            // By default, the data reader is positioned on the first result.
            Reset();
        }

        private object[] CellsValues => _rowIterator?.Current;

        private XlsWorkbook ReadWorkbook(bool convertOADate)
        {
            XlsDirectoryEntry workbookEntry = _document.FindEntry(Workbook) ?? _document.FindEntry(Book);

            if (workbookEntry == null)
            {
                throw new ExcelReaderException(Errors.ErrorStreamWorkbookNotFound);
            }

            if (workbookEntry.EntryType != STGTY.STGTY_STREAM)
            {
                throw new ExcelReaderException(Errors.ErrorWorkbookIsNotStream);
            }

            var bytes = _document.ReadStream(_file, workbookEntry.StreamFirstSector, (int)workbookEntry.StreamSize, workbookEntry.IsEntryMiniStream);

            return new XlsWorkbook(bytes, convertOADate, ReadOption);
        }

        private void ResetSheetData()
        {
            _isFirstRead = true;
            Depth = -1;
        }
    }

    // IDisposable implementation
    public partial class ExcelBinaryReader
    {
        ~ExcelBinaryReader()
        {
            Dispose(false);
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
    }

    // IExcelDataReader implementation.
    public partial class ExcelBinaryReader
    {
        public bool IsFirstRowAsColumnNames { get; set; }

        public ReadOption ReadOption { get; }

        public Encoding Encoding => _workbook.Encoding;

        public string Name => _worksheetIterator?.Current?.Name;

        public string VisibleState => _worksheetIterator?.Current?.VisibleState;

        public int Depth { get; private set; }

        public int ResultsCount => _workbook.Sheets.Count;

        public bool IsClosed { get; private set; }

        public int FieldCount => _worksheetIterator?.Current?.FieldCount ?? 0;

        public int RecordsAffected => throw new NotSupportedException();
        
        public object this[int i] => CellsValues[i];

        public object this[string name] => throw new NotSupportedException();

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

        public string GetDataTypeName(int i)
        {
            throw new NotSupportedException();
        }

        public Guid GetGuid(int i)
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

        public Type GetFieldType(int i)
        {
            return CellsValues[i]?.GetType();
        }

        public string GetName(int i)
        {
            return _cellsNames?[i];
        }

        /// <inheritdoc />
        public IDataReader GetData(int i)
        {
            throw new NotSupportedException();
        }

        /// <inheritdoc />
        public DataTable GetSchemaTable()
        {
            throw new NotSupportedException();
        }
        
        public void Reset()
        {
            _worksheetIterator?.Dispose();
            _rowIterator?.Dispose();

            _worksheetIterator = null;
            _rowIterator = null;

            ResetSheetData();

            if (_workbook != null)
            { 
                _worksheetIterator = _workbook.ReadWorksheets().GetEnumerator();
                if (!_worksheetIterator.MoveNext())
                {
                    _worksheetIterator.Dispose();
                    _worksheetIterator = null;
                    return;
                }

                _rowIterator = _worksheetIterator.Current.ReadRows().GetEnumerator();
            }
        }

        public void Close()
        {
            if (IsClosed)
                return;

            _worksheetIterator?.Dispose();
            _rowIterator?.Dispose();
            _file?.Dispose();

            _worksheetIterator = null;
            _rowIterator = null;
            _workbook = null;
            _document = null;
            _file = null;

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

            if (_isFirstRead)
            {
                _isFirstRead = false;
                if (IsFirstRowAsColumnNames)
                {
                    _cellsNames = new string[CellsValues.Length];
                    for (var i = 0; i < CellsValues.Length; i++)
                    {
                        var value = CellsValues[i]?.ToString();
                        if (!string.IsNullOrEmpty(value))
                        {
                            _cellsNames[i] = value;
                        }
                    }

                    if (!_rowIterator.MoveNext())
                    {
                        _rowIterator.Dispose();
                        _rowIterator = null;
                        return false;
                    }
                }
                else
                {
                    _cellsNames = null; // no columns
                }
            }

            Depth++;
            return true;
        }

        public bool GetBoolean(int i)
        {
            return !IsDBNull(i) && bool.Parse(CellsValues[i].ToString());
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
            return IsDBNull(i) ? decimal.MinValue : decimal.Parse(CellsValues[i].ToString());
        }

        public double GetDouble(int i)
        {
            return IsDBNull(i) ? double.MinValue : double.Parse(CellsValues[i].ToString());
        }

        public float GetFloat(int i)
        {
            return IsDBNull(i) ? float.MinValue : float.Parse(CellsValues[i].ToString());
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

        public string GetString(int i)
        {
            return IsDBNull(i) ? null : CellsValues[i].ToString();
        }

        public object GetValue(int i)
        {
            return CellsValues[i];
        }

        public bool IsDBNull(int i)
        {
            return CellsValues[i] == null;
        }
    }
}