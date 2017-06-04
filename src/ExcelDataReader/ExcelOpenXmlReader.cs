// #define DEBUGREADERS

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using ExcelDataReader.Core;
using ExcelDataReader.Core.OpenXmlFormat;

namespace ExcelDataReader
{
    public partial class ExcelOpenXmlReader : IExcelDataReader
    {
        private const string ElementSheet = "sheet";
        private const string ElementT = "t";
        private const string ElementStringItem = "si";
        private const string ElementCellCrossReference = "cellXfs";
        private const string ElementNumberFormats = "numFmts";

        private const string AttributeSheetId = "sheetId";
        private const string AttributeVisibleState = "state";
        private const string AttributeName = "name";
        private const string AttributeRelationshipId = "r:id";

        private const string ElementRelationship = "Relationship";
        private const string AttributeId = "Id";
        private const string AttributeTarget = "Target";

        private readonly List<int> _defaultDateTimeStyles;

        private XlsxWorkbook _workbook;

        private bool _isFirstRead;

        private int _resultIndex;
        private int _emptyRowCount;
        private ZipWorker _zipWorker;
        private XmlReader _xmlReader;
        private Stream _sheetStream;
        private string[] _cellsNames;
        private object[] _cellsValues;
        private object[] _savedCellsValues;

        private string _namespaceUri;

        public ExcelOpenXmlReader(Stream stream)
        {
            _isFirstRead = true;

            _defaultDateTimeStyles = new List<int>(new[] 
            {
                14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
            });

            _zipWorker = new ZipWorker(stream);

            ReadGlobals();
        }
        
        private static List<XlsxWorksheet> ReadWorkbook(Stream xmlFileStream)
        {
            var sheets = new List<XlsxWorksheet>();

            using (XmlReader reader = XmlReader.Create(xmlFileStream))
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementSheet)
                    {
                        sheets.Add(new XlsxWorksheet(
                            reader.GetAttribute(AttributeName),
                            int.Parse(reader.GetAttribute(AttributeSheetId)),
                            reader.GetAttribute(AttributeRelationshipId),
                            reader.GetAttribute(AttributeVisibleState)));
                    }
                }
            }

            return sheets;
        }

        private static void ReadWorkbookRels(Stream xmlFileStream, List<XlsxWorksheet> sheets)
        {
            using (XmlReader reader = XmlReader.Create(xmlFileStream))
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementRelationship)
                    {
                        string rid = reader.GetAttribute(AttributeId);

                        for (int i = 0; i < sheets.Count; i++)
                        {
                            XlsxWorksheet tempSheet = sheets[i];

                            if (tempSheet.Rid == rid)
                            {
                                tempSheet.Path = reader.GetAttribute(AttributeTarget);
                                sheets[i] = tempSheet;
                                break;
                            }
                        }
                    }
                }
            }
        }

        private static XlsxSST ReadSharedStrings(Stream xmlFileStream)
        {
            if (xmlFileStream == null)
                return null;

            var sst = new XlsxSST();

            using (XmlReader reader = XmlReader.Create(xmlFileStream))
            {
                // Skip phonetic string data.
                bool bSkipPhonetic = false;
                
                // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                bool bAddStringItem = false;
                string sStringItem = string.Empty;

                while (reader.Read())
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementStringItem)
                    {
                        // Do not add the string item until the next string item is read.
                        if (bAddStringItem)
                        {
                            // Add the string item to XlsxSST.
                            sst.Add(sStringItem);
                        }
                        else
                        {
                            // Add the string items from here on.
                            bAddStringItem = true;
                        }

                        // Reset the string item.
                        sStringItem = string.Empty;
                    }
                    else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementT)
                    {
                        // Skip phonetic string data.
                        if (!bSkipPhonetic)
                        {
                            // Append to the string item.
                            sStringItem += reader.ReadElementContentAsString();
                        }
                    }

                    if (reader.LocalName == "rPh")
                    {
                        // Phonetic items represents pronunciation hints for some East Asian languages.
                        // In the file 'xl/sharedStrings.xml', the phonetic properties appear like:
                        // <si>
                        //  <t>(a japanese text in KANJI)</t>
                        //  <rPh sb="0" eb="1">
                        //      <t>(its pronounciation in KATAKANA)</t>
                        //  </rPh>
                        // </si>
                        if (reader.NodeType == XmlNodeType.Element)
                            bSkipPhonetic = true;
                        else if (reader.NodeType == XmlNodeType.EndElement)
                            bSkipPhonetic = false;
                    }
                }

                // Do not add the last string item unless we have read previous string items.
                if (bAddStringItem)
                {
                    // Add the string item to XlsxSST.
                    sst.Add(sStringItem);
                }
            }

            return sst;
        }

        private static XlsxStyles ReadStyles(Stream xmlFileStream)
        {
            var styles = new XlsxStyles();

            if (xmlFileStream == null)
                return styles;

            bool rXlsxNumFmt = false;

            using (XmlReader reader = XmlReader.Create(xmlFileStream))
            {
                while (reader.Read())
                {
                    if (!rXlsxNumFmt && reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementNumberFormats)
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1)
                                break;

                            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxNumFmt.NNumFmt)
                            {
                                styles.NumFmts.Add(
                                    new XlsxNumFmt(
                                        int.Parse(reader.GetAttribute(XlsxNumFmt.ANumFmtId)),
                                        reader.GetAttribute(XlsxNumFmt.AFormatCode)));
                            }
                        }

                        rXlsxNumFmt = true;
                    }

                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementCellCrossReference)
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1)
                                break;

                            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxXf.NXF)
                            {
                                var xfId = reader.GetAttribute(XlsxXf.AXFId);
                                var numFmtId = reader.GetAttribute(XlsxXf.ANumFmtId);

                                styles.CellXfs.Add(
                                    new XlsxXf(
                                        xfId == null ? -1 : int.Parse(xfId),
                                        numFmtId == null ? -1 : int.Parse(numFmtId),
                                        reader.GetAttribute(XlsxXf.AApplyNumberFormat)));
                            }
                        }

                        break;
                    }
                }
            }

            return styles;
        }

        private void ReadGlobals()
        {
            List<XlsxWorksheet> sheets;
            XlsxSST sst;
            XlsxStyles styles;

            using (var stream = _zipWorker.GetWorkbookStream())
            {
                sheets = ReadWorkbook(stream);
            }

            using (var stream = _zipWorker.GetWorkbookRelsStream())
            {
                ReadWorkbookRels(stream, sheets);
            }

            using (var stream = _zipWorker.GetSharedStringsStream())
            {
                sst = ReadSharedStrings(stream);
            }

            using (var stream = _zipWorker.GetStylesStream())
            {
                styles = ReadStyles(stream);
            }

            _workbook = new XlsxWorkbook(sheets, sst, styles);

            CheckDateTimeNumFmts(_workbook.Styles.NumFmts);
        }

        private void CheckDateTimeNumFmts(List<XlsxNumFmt> list)
        {
            if (list.Count == 0)
                return;

            foreach (XlsxNumFmt numFmt in list)
            {
                if (string.IsNullOrEmpty(numFmt.FormatCode))
                    continue;
                string fc = numFmt.FormatCode.ToLower();

                int pos;
                while ((pos = fc.IndexOf('"')) > 0)
                {
                    int endPos = fc.IndexOf('"', pos + 1);

                    if (endPos > 0)
                        fc = fc.Remove(pos, endPos - pos + 1);
                }

                // it should only detect it as a date if it contains
                // dd mm mmm yy yyyy
                // h hh ss
                // AM PM
                // and only if these appear as "words" so either contained in [ ]
                // or delimted in someway
                // updated to not detect as date if format contains a #
                var formatReader = new FormatReader { FormatString = fc };
                if (formatReader.IsDateFormatString())
                {
                    _defaultDateTimeStyles.Add(numFmt.Id);
                }
            }
        }

        private void ReadSheetGlobals(XlsxWorksheet sheet)
        {
            ((IDisposable)_xmlReader)?.Dispose();
            _sheetStream?.Dispose();

            _sheetStream = _zipWorker.GetWorksheetStream(sheet.Path);

            if (_sheetStream == null)
                return;

            _xmlReader = XmlReader.Create(_sheetStream);

            // count rows and cols in case there is no dimension elements
            int rows = 0;
            int cols = 0;

            bool foundDimension = false;

            _namespaceUri = null;
            int biggestColumn = 0; // used when no col elements and no dimension
            int cellElementsInRow = 0;
            while (_xmlReader.Read())
            {
                if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.NWorksheet)
                {
                    // grab the namespaceuri from the worksheet element
                    _namespaceUri = _xmlReader.NamespaceURI;
                }
                
                if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.NDimension)
                {
                    string dimValue = _xmlReader.GetAttribute(XlsxWorksheet.ARef);

                    var dimension = new XlsxDimension(dimValue);
                    if (dimension.IsRange)
                    {
                        sheet.Dimension = dimension;
                        foundDimension = true;

                        break;
                    }
                }

                // removed: Do not use col to work out number of columns as this is really for defining formatting, so may not contain all columns
                /*if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_col)
                    cols++;*/

                if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.NRow)
                {
                    biggestColumn = Math.Max(biggestColumn, cellElementsInRow);
                    cellElementsInRow = 0;
                    rows++;
                }

                // check cells so we can find size of sheet if can't work it out from dimension or col elements (dimension should have been set before the cells if it was available)
                // ditto for cols
                if (cols == 0 && _xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.NC)
                {
                    cellElementsInRow++; 

                    var refAttribute = _xmlReader.GetAttribute(XlsxWorksheet.AR);

                    if (refAttribute != null)
                    {
                        int column;
                        ReferenceHelper.ParseReference(refAttribute, out column);
                        if (column > biggestColumn)
                            biggestColumn = column;
                    }
                }
            }

            biggestColumn = Math.Max(biggestColumn, cellElementsInRow);

            // if we didn't get a dimension element then use the calculated rows/cols to create it
            if (!foundDimension)
            {
                if (cols == 0)
                    cols = biggestColumn;

                if (rows == 0 || cols == 0) 
                {
                    sheet.IsEmpty = true;
                    return;
                }

                sheet.Dimension = new XlsxDimension(rows, cols);

                // we need to reset our position to sheet data
                ((IDisposable)_xmlReader).Dispose();
                _sheetStream.Dispose();
                _sheetStream = _zipWorker.GetWorksheetStream(sheet.Path);
                _xmlReader = XmlReader.Create(_sheetStream);
            }

            // read up to the sheetData element. if this element is empty then there aren't any rows and we need to null out dimension
            _xmlReader.ReadToFollowing(XlsxWorksheet.NSheetData, _namespaceUri);
            if (_xmlReader.IsEmptyElement)
            {
                sheet.IsEmpty = true;
            }
        }

        private bool ReadSheetRow(XlsxWorksheet sheet)
        {
            if (_xmlReader == null)
                return false;

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

            if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.NRow ||
                _xmlReader.ReadToFollowing(XlsxWorksheet.NRow, _namespaceUri))
            {
                _cellsValues = new object[sheet.ColumnsCount];

                int rowIndex;
                if (!int.TryParse(_xmlReader.GetAttribute(XlsxWorksheet.AR), out rowIndex))
                    rowIndex = Depth + 1;

                if (rowIndex != (Depth + 1))
                {
                    _emptyRowCount = rowIndex - Depth - 1;
                }

                bool hasValue = false;
                string aS = string.Empty;
                string aT = string.Empty;
                string aR = string.Empty;
                int col = 0;

                while (_xmlReader.Read())
                {
                    if (_xmlReader.Depth == 2)
                        break;

                    if (_xmlReader.NodeType == XmlNodeType.Element)
                    {
                        hasValue = false;

                        if (_xmlReader.LocalName == XlsxWorksheet.NC)
                        {
                            aS = _xmlReader.GetAttribute(XlsxWorksheet.AS);
                            aT = _xmlReader.GetAttribute(XlsxWorksheet.AT);
                            aR = _xmlReader.GetAttribute(XlsxWorksheet.AR);

                            if (aR != null)
                            {
                                ReferenceHelper.ParseReference(aR, out col);
                            }
                            else
                            {
                                ++col;
                            }
                        }
                        else if (_xmlReader.LocalName == XlsxWorksheet.NV || _xmlReader.LocalName == XlsxWorksheet.NT)
                        {
                            hasValue = true;
                        }
                    }

                    if (_xmlReader.NodeType == XmlNodeType.Text && hasValue)
                    {
                        double number;
                        object o = _xmlReader.Value;

                        const NumberStyles style = NumberStyles.Any;
                        var culture = CultureInfo.InvariantCulture;
                        
                        if (double.TryParse(o.ToString(), style, culture, out number))
                            o = number;
                            
                        if (aT != null && aT == XlsxWorksheet.AS) //// if string
                        {
                            o = Helpers.ConvertEscapeChars(_workbook.SST[int.Parse(o.ToString())]);
                        } // Requested change 4: missing (it appears that if should be else if)
                        else if (aT != null && aT == XlsxWorksheet.NInlineStr) //// if string inline
                        {
                            o = Helpers.ConvertEscapeChars(o.ToString());
                        }
                        else if (aT == "b") //// boolean
                        {
                            o = _xmlReader.Value == "1";
                        }  
                        else if (aS != null) //// if something else
                        {
                            XlsxXf xf = _workbook.Styles.CellXfs[int.Parse(aS)];
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

                Depth++;

                return true;
            }

            ((IDisposable)_xmlReader).Dispose();
            _sheetStream?.Dispose();

            return false;
        }

        private bool InitializeSheetRead()
        {
            if (ResultsCount <= 0)
                return false;

            ReadSheetGlobals(_workbook.Sheets[_resultIndex]);

            if (_workbook.Sheets[_resultIndex].Dimension == null)
                return false;

            _isFirstRead = false;

            Depth = 0;
            _emptyRowCount = 0;

            return true;
        }

        private bool IsDateTimeStyle(int styleId)
        {
            return _defaultDateTimeStyles.Contains(styleId);
        }
    }

    public partial class ExcelOpenXmlReader
    {
        ~ExcelOpenXmlReader()
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

    public partial class ExcelOpenXmlReader
    {
        public bool IsFirstRowAsColumnNames { get; set; }

        public bool ConvertOaDate { get; set; }

        public ReadOption ReadOption { get; set; }

        public Encoding Encoding => null;

        public string Name => _resultIndex >= 0 && _resultIndex < ResultsCount ? _workbook.Sheets[_resultIndex].Name : null;

        public string VisibleState => _resultIndex >= 0 && _resultIndex < ResultsCount ? _workbook.Sheets[_resultIndex].VisibleState : null;

        public int Depth { get; private set; }

        public int ResultsCount => _workbook?.Sheets.Count ?? -1;

        public bool IsClosed { get; private set; }

        public int FieldCount => _resultIndex >= 0 && _resultIndex < ResultsCount ? _workbook.Sheets[_resultIndex].ColumnsCount : -1;

        public int RecordsAffected => throw new NotSupportedException();

        public object this[int i] => _cellsValues[i];

        public object this[string name] => throw new NotSupportedException();

        public void Reset()
        {
            _resultIndex = 0;
            _isFirstRead = true;
            _savedCellsValues = null;
        }

        public void Close()
        {
            if (IsClosed)
                return;

            ((IDisposable)_xmlReader)?.Dispose();
            _sheetStream?.Dispose();
            _zipWorker?.Dispose();

            _zipWorker = null;
            _xmlReader = null;
            _sheetStream = null;

            _workbook = null;
            _cellsValues = null;
            _savedCellsValues = null;

            IsClosed = true;
        }

        public bool NextResult()
        {
            if (_resultIndex >= ResultsCount - 1)
                return false;

            _resultIndex++;

            _isFirstRead = true;
            _savedCellsValues = null;

            return true;
        }

        public bool Read()
        {
            if (_isFirstRead)
            {
                var initializeSheetRead = InitializeSheetRead();
                if (!initializeSheetRead)
                    return false;

                if (IsFirstRowAsColumnNames)
                {
                    if (ReadSheetRow(_workbook.Sheets[_resultIndex]))
                    {
                        _cellsNames = new string[_cellsValues.Length];
                        for (var i = 0; i < _cellsValues.Length; i++)
                        {
                            var value = _cellsValues[i]?.ToString();
                            if (!string.IsNullOrEmpty(value))
                                _cellsNames[i] = value;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    _cellsNames = null;
                }
            }

            return ReadSheetRow(_workbook.Sheets[_resultIndex]);
        }

        public bool GetBoolean(int i)
        {
            return !IsDBNull(i) && bool.Parse(_cellsValues[i].ToString());
        }

        public DateTime GetDateTime(int i)
        {
            if (IsDBNull(i))
                return DateTime.MinValue;

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
            return IsDBNull(i) ? decimal.MinValue : decimal.Parse(_cellsValues[i].ToString());
        }

        public double GetDouble(int i)
        {
            return IsDBNull(i) ? double.MinValue : double.Parse(_cellsValues[i].ToString());
        }

        public float GetFloat(int i)
        {
            return IsDBNull(i) ? float.MinValue : float.Parse(_cellsValues[i].ToString());
        }

        public short GetInt16(int i)
        {
            return IsDBNull(i) ? short.MinValue : short.Parse(_cellsValues[i].ToString());
        }

        public int GetInt32(int i)
        {
            return IsDBNull(i) ? int.MinValue : int.Parse(_cellsValues[i].ToString());
        }

        public long GetInt64(int i)
        {
            return IsDBNull(i) ? long.MinValue : long.Parse(_cellsValues[i].ToString());
        }

        public string GetString(int i)
        {
            return IsDBNull(i) ? null : _cellsValues[i].ToString();
        }

        public object GetValue(int i)
        {
            return _cellsValues[i];
        }

        public bool IsDBNull(int i)
        {
            return _cellsValues[i] == null;
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

        public Type GetFieldType(int i)
        {
            if (_cellsValues[i] == null)
                return null;
            return _cellsValues[i].GetType();
        }

        public Guid GetGuid(int i)
        {
            throw new NotSupportedException();
        }

        public string GetName(int i)
        {
            return _cellsNames?[i];
        }

        public int GetOrdinal(string name)
        {
            throw new NotSupportedException();
        }

        public int GetValues(object[] values)
        {
            throw new NotSupportedException();
        }

        /// <inheritdoc />
        public DataTable GetSchemaTable()
        {
            throw new NotSupportedException();
        }
    }
}
