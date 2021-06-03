using System;
using System.Collections.Generic;
using System.Data;
using ExcelDataReader.Core;

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
        private IEnumerator<Row> _rowIterator;
        private IEnumerator<TWorksheet> _cachedWorksheetIterator;
        private List<TWorksheet> _cachedWorksheets;

        ~ExcelDataReader()
        {
            Dispose(false);
        }

        public string Name => _worksheetIterator?.Current?.Name;

        public string CodeName => _worksheetIterator?.Current?.CodeName;

        public string VisibleState => _worksheetIterator?.Current?.VisibleState;

        public HeaderFooter HeaderFooter => _worksheetIterator?.Current?.HeaderFooter;

        // We shouldn't expose the internal array here. 
        public CellRange[] MergeCells => _worksheetIterator?.Current?.MergeCells;

        public int Depth { get; private set; }

        public int ResultsCount => Workbook?.ResultsCount ?? -1;

        public bool IsClosed { get; private set; }

        public int FieldCount => _worksheetIterator?.Current?.FieldCount ?? 0;

        public int RowCount => _worksheetIterator?.Current?.RowCount ?? 0;

        public int RecordsAffected => throw new NotSupportedException();

        public double RowHeight => _rowIterator?.Current.Height ?? 0;

        protected TWorkbook Workbook { get; set; }

        protected Cell[] RowCells { get; set; }

        public object this[int i] => GetValue(i);

        public object this[string name] => throw new NotSupportedException();

        public bool GetBoolean(int i) => (bool)GetValue(i);

        public byte GetByte(int i) => (byte)GetValue(i);

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
            => throw new NotSupportedException();

        public char GetChar(int i) => (char)GetValue(i);

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
             => throw new NotSupportedException();

        public IDataReader GetData(int i) => throw new NotSupportedException();

        public string GetDataTypeName(int i) => throw new NotSupportedException();

        public DateTime GetDateTime(int i) => (DateTime)GetValue(i);

        public decimal GetDecimal(int i) => (decimal)GetValue(i);

        public double GetDouble(int i) => (double)GetValue(i);

        public Type GetFieldType(int i) => GetValue(i)?.GetType();

        public float GetFloat(int i) => (float)GetValue(i);

        public Guid GetGuid(int i) => (Guid)GetValue(i);

        public short GetInt16(int i) => (short)GetValue(i);

        public int GetInt32(int i) => (int)GetValue(i);

        public long GetInt64(int i) => (long)GetValue(i);

        public string GetName(int i)
        {
            return FieldNames()[i];
        }

        private string[] _fieldNames = null;
        private Cell[] _fieldNameCells = null;

        public Dictionary<string, object> InjectedFields { get; set; } // optionally set in ExcelOpenXmlReader, default to null when not specified

        /// <summary>
        /// return an array of field names found in the first line of the xls; reads and does NOT reset row counter to then skip field names during stream bulk insert
        /// </summary>
        /// <returns>array of field names</returns>
        public string[] FieldNames()
        {
            if (_fieldNames == null)
            {
                // the first time FieldNames() gets called, the list will need to be initialized
                if (RowCells == null)
                {
                    Read();
                    _fieldNameCells = RowCells;

                    // Reset(); // the first row *must* contain field names for bulk copy streams. When assigning the field names here, we do not reset the row pointer.
                }
                else
                {
                    _fieldNameCells = RowCells;
                }

                if (InjectedFields == null)
                {
                    // no InjectedFields found. The length of out field name array will be the same as the number of RowCells
                    _fieldNames = new string[_fieldNameCells.Length];
                }
                else
                {
                    // InjectedFields specified; our total field list size will include the RowCells, plus the number of injected fields 
                    _fieldNames = new string[_fieldNameCells.Length + InjectedFields.Count];

                    // if there are any static fields to be injected, add them to the end of the field list
                    int i = 0;
                    foreach (string injectedFieldName in InjectedFields.Keys)
                    {
                        _fieldNames[_fieldNameCells.Length + i] = injectedFieldName;
                        i++;
                    }
                }

                // in ether case, InjectedFields or not ... save all the field names from the first row of RowCells for future reference
                for (int i = 0; i < _fieldNameCells.Length; i++)
                {
                    _fieldNames[i] = (string)_fieldNameCells[i].Value;
                }
            }

            return _fieldNames;
        }

        /// <summary>
        /// typically only used by bulk insert column mapping, GetOrdinal returns a column number
        /// </summary>
        /// <param name="name">column name</param>
        /// <returns>column number</returns>
        public int GetOrdinal(string name)
        {
            return Array.IndexOf(FieldNames(), name);
        }

        /// <inheritdoc />
        public DataTable GetSchemaTable() => throw new NotSupportedException();

        public string GetString(int i) => (string)GetValue(i);

        /// <summary>
        /// return the value of a cell at ordinal value [i]; may also return a fixed value as defined in 
        /// </summary>
        /// <param name="i">ordinal column referenbce</param>
        /// <returns>value of that cell, or null</returns>
        public object GetValue(int i)
        {
            if (RowCells == null)
                throw new InvalidOperationException("No data exists for the row/column.");

            // if fixed columns values were injected then FieldNames().Length > RowCells.Length; Recall Length is 1-based
            if (i >= RowCells.Length)
            {
                if (_fieldNames == null)
                {
                    throw new Exception("ERROR: _fieldNames unexpectedly empty in ExcelDataReader.");
                }

                if ((_fieldNames.Length - i) > InjectedFields.Count)
                {
                    throw new Exception("ERROR: Ordinal value of " + i.ToString() + " exceeds number of available InjectedFields static values provided to ExcelDataReader.");
                }
            }

            string thisField = GetName(i);

            if (InjectedFields != null && InjectedFields.ContainsKey(thisField))
            {
                return InjectedFields[thisField];
            }
            else
            {
                return RowCells[i]?.Value; // RowCells is zero-based
            }
        }

        public int GetValues(object[] values) => throw new NotSupportedException();

        public bool IsDBNull(int i) => GetValue(i) == null;

        public string GetNumberFormatString(int i)
        {
            if (RowCells == null)
                throw new InvalidOperationException("No data exists for the row/column.");
            if (RowCells[i] == null)
                return null;
            return _worksheetIterator?.Current?.GetNumberFormatString(RowCells[i].NumberFormatIndex)?.FormatString;
        }

        public int GetNumberFormatIndex(int i)
        {
            if (RowCells == null)
                throw new InvalidOperationException("No data exists for the row/column.");
            if (RowCells[i] == null)
                return -1;
            return RowCells[i].NumberFormatIndex;
        }

        public double GetColumnWidth(int i)
        {
            if (i >= FieldCount)
            {
                throw new ArgumentException($"Column at index {i} does not exist.", nameof(i));
            }

            var columnWidths = _worksheetIterator?.Current?.ColumnWidths ?? null;
            double? retWidth = null;
            if (columnWidths != null)
            {
                var colWidthIndex = 0;
                while (colWidthIndex < columnWidths.Length && retWidth == null)
                {
                    var columnWidth = columnWidths[colWidthIndex];
                    if (i >= columnWidth.Min && i <= columnWidth.Max)
                    {
                        retWidth = columnWidth.Hidden
                            ? 0
                            : columnWidth.Width;
                    }
                    else
                    {
                        colWidthIndex++;
                    }
                }
            }

            const double DefaultColumnWidth = 8.43D;

            return retWidth ?? DefaultColumnWidth;
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
                _worksheetIterator = ReadWorksheetsWithCache().GetEnumerator(); // Workbook.ReadWorksheets().GetEnumerator();
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
            RowCells = null;
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

            ReadCurrentRow();

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

        private IEnumerable<TWorksheet> ReadWorksheetsWithCache()
        {
            // Iterate TWorkbook.ReadWorksheets() only once and cache the 
            // worksheet instances, which are expensive to create. 
            if (_cachedWorksheets != null)
            {
                foreach (var worksheet in _cachedWorksheets)
                {
                    yield return worksheet;
                }

                if (_cachedWorksheetIterator == null)
                {
                    yield break;
                }
            }
            else
            {
                _cachedWorksheets = new List<TWorksheet>();
            }

            if (_cachedWorksheetIterator == null)
            {
                _cachedWorksheetIterator = Workbook.ReadWorksheets().GetEnumerator();
            }

            while (_cachedWorksheetIterator.MoveNext())
            {
                _cachedWorksheets.Add(_cachedWorksheetIterator.Current);
                yield return _cachedWorksheetIterator.Current;
            }

            _cachedWorksheetIterator.Dispose();
            _cachedWorksheetIterator = null;
        }

        private void ResetSheetData()
        {
            Depth = -1;
            RowCells = null;
        }

        private void ReadCurrentRow()
        {
            if (RowCells == null)
            {
                RowCells = new Cell[FieldCount];
            }

            Array.Clear(RowCells, 0, RowCells.Length);

            foreach (var cell in _rowIterator.Current.Cells)
            {
                if (cell.ColumnIndex < RowCells.Length)
                {
                    RowCells[cell.ColumnIndex] = cell;
                }
            }
        }
    }
}
