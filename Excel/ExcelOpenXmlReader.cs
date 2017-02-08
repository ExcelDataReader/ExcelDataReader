#define DEBUGREADERS

using System;
using System.Collections.Generic;
using System.Text;
using Excel.Core.OpenXmlFormat;
using System.IO;
using System.Text;
using Excel.Core;
using System.Data;
using System.Xml;
using System.Globalization;

namespace Excel
{

    public class ExcelOpenXmlReader : IExcelDataReader
    {
        #region Members

        private XlsxWorkbook _workbook;
        private bool _isValid;
        private bool _isClosed;
        private bool _isFirstRead;
        private string _exceptionMessage;
        private int _depth;
        private int _resultIndex;
        private int _emptyRowCount;
        private ZipWorker _zipWorker;
        private XmlReader _xmlReader;
        private Stream _sheetStream;
        private object[] _cellsValues;
        private object[] _savedCellsValues;

        private bool disposed;
        private bool _isFirstRowAsColumnNames;
        private const string COLUMN = "Column";
        private string instanceId = Guid.NewGuid().ToString();

        private List<int> _defaultDateTimeStyles;
        private string _namespaceUri;

        #region Fields for batch support
        private DataSet _schema = null;
        private int _batchSize = 1000;
        private int _sheetIndex = -1;
        private string _sheetName = string.Empty;
        private int _skipRows = -1;
        private DataTable _dtBatch = null;
        #endregion

        #endregion

        internal ExcelOpenXmlReader()
        {
            _isValid = true;
            _isFirstRead = true;

            _defaultDateTimeStyles = new List<int>(new int[]
            {
                14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
            });

        }

        private void ReadGlobals()
        {
            _workbook = new XlsxWorkbook(
                _zipWorker.GetWorkbookStream(),
                _zipWorker.GetWorkbookRelsStream(),
                _zipWorker.GetSharedStringsStream(),
                _zipWorker.GetStylesStream());

            CheckDateTimeNumFmts(_workbook.Styles.NumFmts);

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
                    _defaultDateTimeStyles.Add(numFmt.Id);
                }
            }
        }

        private void ReadSheetGlobals(XlsxWorksheet sheet)
        {
            if (_xmlReader != null) _xmlReader.Close();
            if (_sheetStream != null) _sheetStream.Close();

            _sheetStream = _zipWorker.GetWorksheetStream(sheet.Path);

            if (null == _sheetStream) return;

            _xmlReader = XmlReader.Create(_sheetStream);

            //count rows and cols in case there is no dimension elements
            int rows = 0;
            int cols = 0;
            sheet.Dimension = null;
            _namespaceUri = null;
            int biggestColumn = 0; //used when no col elements and no dimension
            while (_xmlReader.Read())
            {
                if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_worksheet)
                {
                    //grab the namespaceuri from the worksheet element
                    _namespaceUri = _xmlReader.NamespaceURI;
                }

                if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_dimension)
                {
                    string dimValue = _xmlReader.GetAttribute(XlsxWorksheet.A_ref);
                    // To fix the issue of loading kendo generated excel files.
                    if (dimValue.IndexOf(':') > -1)
                    {
                        sheet.Dimension = new XlsxDimension(dimValue);
                        break;
                    }
                }

                //removed: Do not use col to work out number of columns as this is really for defining formatting, so may not contain all columns
                //if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_col)
                //    cols++;

                if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_row)
                    rows++;

                //check cells so we can find size of sheet if can't work it out from dimension or col elements (dimension should have been set before the cells if it was available)
                //ditto for cols
                if (sheet.Dimension == null && cols == 0 && _xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_c)
                {
                    var refAttribute = _xmlReader.GetAttribute(XlsxWorksheet.A_r);

                    if (refAttribute != null)
                    {
                        var thisRef = ReferenceHelper.ReferenceToColumnAndRow(refAttribute);
                        if (thisRef[1] > biggestColumn)
                            biggestColumn = thisRef[1];
                    }
                }

            }


            //if we didn't get a dimension element then use the calculated rows/cols to create it
            if (sheet.Dimension == null)
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
                _xmlReader.Close();
                _sheetStream.Close();
                _sheetStream = _zipWorker.GetWorksheetStream(sheet.Path);
                _xmlReader = XmlReader.Create(_sheetStream);

            }

            //read up to the sheetData element. if this element is empty then there aren't any rows and we need to null out dimension

            _xmlReader.ReadToFollowing(XlsxWorksheet.N_sheetData, _namespaceUri);
            if (_xmlReader.IsEmptyElement)
            {
                sheet.IsEmpty = true;
            }


        }

        private bool ReadSheetRow(XlsxWorksheet sheet)
        {
            if (null == _xmlReader) return false;

            if (_emptyRowCount != 0)
            {
                _cellsValues = new object[sheet.ColumnsCount];
                _emptyRowCount--;
                _depth++;

                return true;
            }

            if (_savedCellsValues != null)
            {
                _cellsValues = _savedCellsValues;
                _savedCellsValues = null;
                _depth++;

                return true;
            }

            if ((_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_row) ||
                _xmlReader.ReadToFollowing(XlsxWorksheet.N_row, _namespaceUri))
            {
                _cellsValues = new object[sheet.ColumnsCount];

                int rowIndex = int.Parse(_xmlReader.GetAttribute(XlsxWorksheet.A_r));
                if (rowIndex != (_depth + 1))
                    if (rowIndex != (_depth + 1))
                    {
                        _emptyRowCount = rowIndex - _depth - 1;
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

                        if (_xmlReader.LocalName == XlsxWorksheet.N_c)
                        {
                            a_s = _xmlReader.GetAttribute(XlsxWorksheet.A_s);
                            a_t = _xmlReader.GetAttribute(XlsxWorksheet.A_t);
                            a_r = _xmlReader.GetAttribute(XlsxWorksheet.A_r);
                            XlsxDimension.XlsxDim(a_r, out col, out row);
                        }
                        else if (_xmlReader.LocalName == XlsxWorksheet.N_v || _xmlReader.LocalName == XlsxWorksheet.N_t)
                        {
                            hasValue = true;
                        }
                    }

                    if (_xmlReader.NodeType == XmlNodeType.Text && hasValue)
                    {
                        double number;
                        object o = _xmlReader.Value;

                        var style = NumberStyles.Any;
                        var culture = CultureInfo.InvariantCulture;

                        if (double.TryParse(o.ToString(), style, culture, out number))
                            o = number;

                        if (null != a_t && a_t == XlsxWorksheet.A_s) //if string
                        {
                            o = Helpers.ConvertEscapeChars(_workbook.SST[int.Parse(o.ToString())]);
                        } // Requested change 4: missing (it appears that if should be else if)
                        else if (null != a_t && a_t == XlsxWorksheet.N_inlineStr) //if string inline
                        {
                            o = Helpers.ConvertEscapeChars(o.ToString());
                        }
                        else if (a_t == "b") //boolean
                        {
                            o = _xmlReader.Value == "1";
                        }
                        else if (null != a_s) //if something else
                        {
                            XlsxXf xf = _workbook.Styles.CellXfs[int.Parse(a_s)];
                            if (o != null && o.ToString() != string.Empty && IsDateTimeStyle(xf.NumFmtId))
                                o = Helpers.ConvertFromOATime(number);
                            else if (xf.NumFmtId == 49)
                                o = o.ToString();
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
                _depth++;

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

            _depth = 0;
            _emptyRowCount = 0;

            return true;
        }

        private bool IsDateTimeStyle(int styleId)
        {
            return _defaultDateTimeStyles.Contains(styleId);
        }


        #region IExcelDataReader Members

        public void Initialize(System.IO.Stream fileStream)
        {
            _zipWorker = new ZipWorker();
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
            return AsDataSet(true);
        }

        public System.Data.DataSet AsDataSet(bool convertOADateTime)
        {
            if (!_isValid) return null;

            DataSet dataset = new DataSet();

            for (int ind = 0; ind < _workbook.Sheets.Count; ind++)
            {
                DataTable table = new DataTable(_workbook.Sheets[ind].Name);

                table.ExtendedProperties.Add("visiblestate", _workbook.Sheets[ind].VisibleState);

                ReadSheetGlobals(_workbook.Sheets[ind]);

                if (_workbook.Sheets[ind].Dimension == null) continue;

                _depth = 0;
                _emptyRowCount = 0;

                //DataTable columns
                if (!_isFirstRowAsColumnNames)
                {
                    for (int i = 0; i < _workbook.Sheets[ind].ColumnsCount; i++)
                    {
                        table.Columns.Add(null, typeof(Object));
                    }
                }
                else if (ReadSheetRow(_workbook.Sheets[ind]))
                {
                    for (int index = 0; index < _cellsValues.Length; index++)
                    {
                        if (_cellsValues[index] != null && _cellsValues[index].ToString().Length > 0)
                            Helpers.AddColumnHandleDuplicate(table, _cellsValues[index].ToString());
                        else
                            Helpers.AddColumnHandleDuplicate(table, string.Concat(COLUMN, index));
                    }
                }
                else continue;

                table.BeginLoadData();

                while (ReadSheetRow(_workbook.Sheets[ind]))
                {
                    table.Rows.Add(_cellsValues);
                }

                if (table.Rows.Count > 0)
                    dataset.Tables.Add(table);
                table.EndLoadData();
            }
            dataset.AcceptChanges();
            Helpers.FixDataTypes(dataset);
            return dataset;
        }

        public bool IsFirstRowAsColumnNames
        {
            get
            {
                return _isFirstRowAsColumnNames;
            }
            set
            {
                _isFirstRowAsColumnNames = value;
            }
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

        public string VisibleState
        {
            get
            {
                return (_resultIndex >= 0 && _resultIndex < ResultsCount) ? _workbook.Sheets[_resultIndex].VisibleState : null;
            }
        }

        public void Close()
        {
            _isClosed = true;

            if (_xmlReader != null) _xmlReader.Close();

            if (_sheetStream != null) _sheetStream.Close();

            if (_zipWorker != null) _zipWorker.Dispose();
        }

        public int Depth
        {
            get { return _depth; }
        }

        public int ResultsCount
        {
            get { return _workbook == null ? -1 : _workbook.Sheets.Count; }
        }

        public bool IsClosed
        {
            get { return _isClosed; }
        }

        public bool NextResult()
        {
            if (_resultIndex >= (this.ResultsCount - 1)) return false;

            _resultIndex++;

            _isFirstRead = true;
            _savedCellsValues = null;

            return true;
        }

        public bool Read()
        {
            if (!_isValid) return false;

            if (_isFirstRead && !InitializeSheetRead())
            {
                return false;
            }

            return ReadSheetRow(_workbook.Sheets[_resultIndex]);
        }

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

        #region Methods for batch support
        private DataTable CreateTableSchema(int sheetIndex)
        {
            XlsxWorksheet worksheet = _workbook.Sheets[sheetIndex];
            DataTable dataTable = new DataTable(worksheet.Name);

            dataTable.ExtendedProperties.Add("visiblestate", worksheet.VisibleState);
            dataTable.ExtendedProperties.Add("SkipRows", _skipRows);
            dataTable.ExtendedProperties.Add("IsFirstRowAsColumnNames", _isFirstRowAsColumnNames);

            //DataTable columns
            if (!_isFirstRowAsColumnNames)
            {
                for (int i = 0, columns = worksheet.ColumnsCount; i < columns; i++)
                {
                    dataTable.Columns.Add(null, typeof(Object));
                }
            }
            else if (_cellsValues != null)
            {
                for (int index = 0; index < _cellsValues.Length; index++)
                {
                    object cellValues = _cellsValues[index];
                    if (cellValues != null && cellValues.ToString().Length > 0)
                        Helpers.AddColumnHandleDuplicate(dataTable, cellValues.ToString());
                    else
                        Helpers.AddColumnHandleDuplicate(dataTable, string.Concat(COLUMN, index));
                }
            }
            return dataTable;
        }
        private bool IsSheetIndexValid(int sheetIndex)
        {
            return (_sheetIndex >= 0 && _sheetIndex < _workbook.Sheets.Count);
        }
        private bool ValidateSheetParameters()
        {
            if (string.IsNullOrEmpty(_sheetName)) { throw new Exception("SheetName property of IExcelDataReader is not set."); }
            else { _sheetIndex = _workbook.Sheets.FindIndex(s => s.Name.ToLower() == _sheetName.ToLower()); }
            if (!IsSheetIndexValid(_sheetIndex)) { throw new Exception(string.Format(@"Sheet '{0}' not found in excel.", _sheetName)); }

            return true;
        }
        private bool InitializeSheetBatchRead()
        {
            if (ResultsCount <= 0) return false;
            XlsxWorksheet worksheet = _workbook.Sheets[_sheetIndex];
            ReadSheetGlobals(worksheet);
            if (worksheet.Dimension == null) return false;
            if (_skipRows >= worksheet.RowsCount) { throw new Exception(string.Format(@"Sheet '{0}' contains less rows than SkipRows property of IExcelDataReader.", _sheetName)); }

            _depth = 0;
            _emptyRowCount = 0;

            if (_skipRows > 0)
            {
                _xmlReader.ReadToFollowing(XlsxWorksheet.N_row, _namespaceUri);
                for (int i = 0; i < _skipRows; i++) { _xmlReader.Skip(); }
                _depth = _skipRows;
            }
            if (_isFirstRowAsColumnNames) { ReadSheetBatchRow(worksheet); }

            _isFirstRead = false;
            return true;
        }

        private bool ReadSheetBatchRow(XlsxWorksheet sheet)
        {
            if (null == _xmlReader) return false;

            if (_emptyRowCount != 0)
            {
                _cellsValues = new object[sheet.ColumnsCount];
                _emptyRowCount--;
                _depth++;

                return true;
            }

            if (_savedCellsValues != null)
            {
                _cellsValues = _savedCellsValues;
                _savedCellsValues = null;
                _depth++;

                return true;
            }

            if ((_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_row) ||
                _xmlReader.ReadToFollowing(XlsxWorksheet.N_row, _namespaceUri))
            {
                _cellsValues = new object[sheet.ColumnsCount];

                int rowIndex = int.Parse(_xmlReader.GetAttribute(XlsxWorksheet.A_r));
                if (rowIndex != (_depth + 1))
                    if (rowIndex != (_depth + 1))
                    {
                        _emptyRowCount = rowIndex - _depth - 1;
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

                        if (_xmlReader.LocalName == XlsxWorksheet.N_c)
                        {
                            a_s = _xmlReader.GetAttribute(XlsxWorksheet.A_s);
                            a_t = _xmlReader.GetAttribute(XlsxWorksheet.A_t);
                            a_r = _xmlReader.GetAttribute(XlsxWorksheet.A_r);
                            XlsxDimension.XlsxDim(a_r, out col, out row);
                        }
                        else if (_xmlReader.LocalName == XlsxWorksheet.N_v || _xmlReader.LocalName == XlsxWorksheet.N_t)
                        {
                            hasValue = true;
                        }
                    }

                    if (_xmlReader.NodeType == XmlNodeType.Text && hasValue)
                    {
                        double number;
                        object o = _xmlReader.Value;

                        var style = NumberStyles.Any;
                        var culture = CultureInfo.InvariantCulture;

                        if (double.TryParse(o.ToString(), style, culture, out number))
                            o = number;

                        if (null != a_t && a_t == XlsxWorksheet.A_s) //if string
                        {
                            o = Helpers.ConvertEscapeChars(_workbook.SST[int.Parse(o.ToString())]);
                        } // Requested change 4: missing (it appears that if should be else if)
                        else if (null != a_t && a_t == XlsxWorksheet.N_inlineStr) //if string inline
                        {
                            o = Helpers.ConvertEscapeChars(o.ToString());
                        }
                        else if (a_t == "b") //boolean
                        {
                            o = _xmlReader.Value == "1";
                        }
                        else if (null != a_s) //if something else
                        {
                            XlsxXf xf = _workbook.Styles.CellXfs[int.Parse(a_s)];
                            if (o != null && o.ToString() != string.Empty && IsDateTimeStyle(xf.NumFmtId))
                                o = Helpers.ConvertFromOATime(number);
                            else if (xf.NumFmtId == 49)
                                o = o.ToString();
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
                _depth++;

                return true;
            }

            _xmlReader.Close();
            if (_sheetStream != null) _sheetStream.Close();

            return false;
        }
        private bool ReadBatchForSchema()
        {
            if (!_isValid) { return false; }

            if (_isFirstRead)
            {
                if (!ValidateSheetParameters()) { return false; }
                if (!InitializeSheetBatchRead()) { return false; }
            }

            _dtBatch = CreateTableSchema(_sheetIndex);
            var workSheet = _workbook.Sheets[_sheetIndex];
            _dtBatch.BeginLoadData();
            int currentRowIndexInBatch = 0;
            while (currentRowIndexInBatch < _batchSize)
            {
                if (ReadSheetBatchRow(workSheet))
                {
                    _dtBatch.LoadDataRow(_cellsValues, false);
                    currentRowIndexInBatch++;
                }
                else { break; }
            }
            _dtBatch.EndLoadData();

            if (currentRowIndexInBatch == 0) { SheetName = string.Empty; return false; }

            return true;
        }

        public int BatchSize
        {
            get
            {
                return _batchSize;
            }

            set
            {
                _batchSize = value;
            }
        }
        public int SheetIndex
        {
            get
            {
                return _sheetIndex;
            }

            set
            {
                _sheetIndex = value;

                _dtBatch = null;
                _isFirstRead = true;
                _savedCellsValues = null;
            }
        }
        public string SheetName
        {
            get
            {
                return _sheetName;
            }

            set
            {
                _sheetName = value;

                _dtBatch = null;
                _isFirstRead = true;               
                _savedCellsValues = null;
            }
        }

        public int SkipRows
        {
            get
            {
                return _skipRows;
            }

            set
            {
                _skipRows = value;
            }
        }


        public List<string> GetSheetNames()
        {
            List<string> sheetNames = new List<string>();
            foreach (var sheet in _workbook.Sheets) { sheetNames.Add(sheet.Name); }
            return sheetNames;
        }

        public DataSet GetSchema(bool isFirstRowAsColumnNames = true, int skipRows = 0)
        {
            if (!_isValid) return null;

            DataSet dataset = new DataSet();
            foreach (var sheet in _workbook.Sheets)
            {
                SheetName = sheet.Name;
                SkipRows = skipRows;
                IsFirstRowAsColumnNames = isFirstRowAsColumnNames;

                if (ReadBatchForSchema()) { dataset.Tables.Add(GetCurrentBatch()); }
            }
            if (dataset.Tables.Count > 0)
            {
                Helpers.FixDataTypes(dataset);
                dataset.Clear();
                _schema = dataset.Clone();
            }
            return dataset;
        }

        public DataSet GetSchema(List<SheetParameters> sheetParametersList)
        {
            if (!_isValid) return null;
            if (sheetParametersList == null || sheetParametersList.Count == 0) return null;

            DataSet dataset = new DataSet();
            foreach (SheetParameters param in sheetParametersList)
            {
                SheetName = param.SheetName;
                SkipRows = param.SkipRows;
                IsFirstRowAsColumnNames = param.IsFirstRowAsColumnNames;

                if (ReadBatchForSchema()) { dataset.Tables.Add(GetCurrentBatch()); }
            }
            if (dataset.Tables.Count > 0)
            {
                Helpers.FixDataTypes(dataset);
                dataset.Clear();
                _schema = dataset.Clone();
            }
            return dataset;
        }

        public DataTable GetSchema(SheetParameters sheetParameters)
        {
            if (!_isValid) return null;

            DataTable dataTable = new DataTable();

            SheetName = sheetParameters.SheetName;
            SkipRows = sheetParameters.SkipRows;
            IsFirstRowAsColumnNames = sheetParameters.IsFirstRowAsColumnNames;

            if (ReadBatchForSchema()) { dataTable = GetCurrentBatch(); }

            if (dataTable.Rows.Count > 0)
            {
                dataTable = Helpers.FixDataTypes(dataTable);
                dataTable.Clear();
                _schema = new DataSet();
                _schema.Tables.Add(dataTable.Clone());
            }
            return dataTable;
        }
        public bool ReadBatch()
        {
            if (!_isValid) { return false; }

            if (_isFirstRead)
            {
                if (!ValidateSheetParameters()) { return false; }
                if (_schema == null || _schema.Tables[_sheetName] == null)
                { GetSchema(new SheetParameters(SheetName, IsFirstRowAsColumnNames, SkipRows)); }

                if (_schema == null || _schema.Tables[_sheetName] == null) { return false; }
                if (!InitializeSheetBatchRead()) { return false; }
                _dtBatch = _schema.Tables[_sheetName].Clone();
            }

            _dtBatch.Clear();

            int currentRowIndexInBatch = 0;
            _dtBatch.BeginLoadData();
            var workSheet = _workbook.Sheets[_sheetIndex];
            while (currentRowIndexInBatch < _batchSize)
            {
                if (ReadSheetBatchRow(workSheet))
                {
                    _dtBatch.LoadDataRow(_cellsValues, false);
                    currentRowIndexInBatch++;
                }
                else { break; }
            }
            _dtBatch.EndLoadData();

            if (currentRowIndexInBatch == 0) { SheetName = string.Empty; return false; }

            return true;
        }
        public DataTable GetCurrentBatch()
        {
            return _dtBatch;
        }

        public DataTable GetTopRows(int rowCount, SheetParameters sheetParameters)
        {
            if (!_isValid) return null;

            DataSet dataset = new DataSet();
            DataTable datatable = null;

            SheetName = sheetParameters.SheetName;
            SkipRows = sheetParameters.SkipRows;
            IsFirstRowAsColumnNames = sheetParameters.IsFirstRowAsColumnNames;

            int originalBatchSize = _batchSize;
            _batchSize = rowCount;
            if (ReadBatchForSchema()) { dataset.Tables.Add(GetCurrentBatch()); }
            _batchSize = originalBatchSize;

            if (dataset.Tables.Count > 0)
            {
                Helpers.FixDataTypes(dataset);
                datatable = dataset.Tables[0];
            }

            return datatable;
        }
        #endregion

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
                    if (_xmlReader != null) ((IDisposable) _xmlReader).Dispose();
                    if (_sheetStream != null) _sheetStream.Dispose();
                    if (_zipWorker != null) _zipWorker.Dispose();
                }

                _zipWorker = null;
                _xmlReader = null;
                _sheetStream = null;

                _workbook = null;
                _cellsValues = null;
                _savedCellsValues = null;

                disposed = true;
            }
        }

        ~ExcelOpenXmlReader()
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


    }
}
