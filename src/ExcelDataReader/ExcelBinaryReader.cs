using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using ExcelDataReader.Core;
using ExcelDataReader.Core.BinaryFormat;
using ExcelDataReader.Exceptions;
using ExcelDataReader.Log;
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
        /* private const string COLUMN = "Column"; */

        private readonly Dictionary<int, object[]> _currentRows = new Dictionary<int, object[]>();

        private Stream _file;
        private XlsHeader _hdr;
        private List<XlsWorksheet> _sheets;
        private XlsBiffStream _stream;
        private XlsWorkbookGlobals _globals;
        private ushort _version;

        private string[] _cellsNames;
        private object[] _cellsValues;
        private int _sheetIndex;

        private bool _isFirstRead;
        private int _largestObservedRow = -1;

        private bool _reachedEndOfSheet;

        public ExcelBinaryReader(Stream stream)
            : this(stream, true, ReadOption.Strict)
        {
        }

        public ExcelBinaryReader(Stream stream, ReadOption readOption)
            : this(stream, true, readOption)
        {
        }

        public ExcelBinaryReader(Stream stream, bool convertOADate, ReadOption readOption)
            : this(stream, new XlsHeader(stream), convertOADate, readOption)
        {
        }

        internal ExcelBinaryReader(Stream stream, XlsHeader header, bool convertOADate, ReadOption readOption)
        {
            _version = 0x0600;
            _isFirstRead = true;
            _file = stream;
            ReadOption = readOption;
            ConvertOaDate = convertOADate;

            _hdr = header;

            if (header.IsRawBiffStream)
                throw new NotSupportedException("File appears to be a raw BIFF stream which isn't supported (BIFF" + header.RawBiffVersion + ").");
            if (!_hdr.IsSignatureValid)
                throw new HeaderException(Errors.ErrorHeaderSignature);
            if (_hdr.ByteOrder != 0xFFFE && _hdr.ByteOrder != 0xFFFF) // Some broken xls files uses 0xFFFF
                throw new FormatException(Errors.ErrorHeaderOrder);

            ReadWorkBookGlobals();

            // set the sheet index to the index of the first sheet.. this is so that properties such as Name which use sheetIndex reflect the first sheet in the file without having to perform a read() operation
            _sheetIndex = 0;
        }

        private bool ConvertOaDate { get; }

        public bool IsV8()
        {
            return _version >= 0x600;
        }

        private void ReadWorkBookGlobals()
        {
            XlsRootDirectory dir = new XlsRootDirectory(_hdr);
            XlsDirectoryEntry workbookEntry = dir.FindEntry(Workbook) ?? dir.FindEntry(Book);

            if (workbookEntry == null)
            {
                throw new ExcelReaderException(Errors.ErrorStreamWorkbookNotFound);
            }

            if (workbookEntry.EntryType != STGTY.STGTY_STREAM)
            {
                throw new ExcelReaderException(Errors.ErrorWorkbookIsNotStream);
            }

            _stream = new XlsBiffStream(_hdr, workbookEntry.StreamFirstSector, workbookEntry.IsEntryMiniStream, dir, this);

            _globals = new XlsWorkbookGlobals();

            _stream.Seek(0, SeekOrigin.Begin);

            XlsBiffRecord rec = _stream.Read();
            XlsBiffBOF bof = rec as XlsBiffBOF;

            if (bof == null || bof.Type != BIFFTYPE.WorkbookGlobals)
            {
                throw new ExcelReaderException(Errors.ErrorWorkbookGlobalsInvalidData);
            }

            bool sst = false;

            _version = bof.Version;
            _sheets = new List<XlsWorksheet>();

            while ((rec = _stream.Read()) != null)
            {
                switch (rec.Id)
                {
                    case BIFFRECORDTYPE.INTERFACEHDR:
                        _globals.InterfaceHdr = (XlsBiffInterfaceHdr)rec;
                        break;
                    case BIFFRECORDTYPE.BOUNDSHEET:
                        XlsBiffBoundSheet sheet = (XlsBiffBoundSheet)rec;

                        if (sheet.Type != XlsBiffBoundSheet.SheetType.Worksheet)
                            break;

                        // sheet.UseEncoding = Encoding;
                        LogManager.Log(this).Debug("BOUNDSHEET IsV8={0}", sheet.IsV8);

                        _sheets.Add(new XlsWorksheet(_globals.Sheets.Count, sheet));
                        _globals.Sheets.Add(sheet);

                        break;
                    case BIFFRECORDTYPE.MMS:
                        _globals.Mms = rec;
                        break;
                    case BIFFRECORDTYPE.COUNTRY:
                        _globals.Country = rec;
                        break;
                    case BIFFRECORDTYPE.CODEPAGE:
                        _globals.CodePage = (XlsBiffSimpleValueRecord)rec;

                        // set encoding based on code page name
                        // PCL does not supported codepage numbers
                        if (_globals.CodePage.Value == 1200)
                            Encoding = EncodingHelper.GetEncoding(65001);
                        else
                            Encoding = EncodingHelper.GetEncoding(_globals.CodePage.Value);

                        // NOTE: the format spec states that for BIFF8 this is always UTF-16.
                        break;
                    case BIFFRECORDTYPE.FONT:
                    case BIFFRECORDTYPE.FONT_V34:
                        _globals.Fonts.Add(rec);
                        break;
                    case BIFFRECORDTYPE.FORMAT_V23:
                        {
                            var fmt = (XlsBiffFormatString)rec;
                            _globals.Formats.Add((ushort)_globals.Formats.Count, fmt);
                        }

                        break;
                    case BIFFRECORDTYPE.FORMAT:
                        {
                            var fmt = (XlsBiffFormatString)rec;
                            _globals.Formats.Add(fmt.Index, fmt);
                        }

                        break;
                    case BIFFRECORDTYPE.XF:
                    case BIFFRECORDTYPE.XF_V4:
                    case BIFFRECORDTYPE.XF_V3:
                    case BIFFRECORDTYPE.XF_V2:
                        _globals.ExtendedFormats.Add(rec);
                        break;
                    case BIFFRECORDTYPE.SST:
                        _globals.SST = (XlsBiffSST)rec;
                        sst = true;
                        break;
                    case BIFFRECORDTYPE.CONTINUE:
                        if (!sst)
                            break;
                        XlsBiffContinue contSST = (XlsBiffContinue)rec;
                        _globals.SST.Append(contSST);
                        break;
                    case BIFFRECORDTYPE.EXTSST:
                        _globals.ExtSST = rec;
                        sst = false;
                        break;
                    case BIFFRECORDTYPE.PASSWORD:
                        break;
                    case BIFFRECORDTYPE.PROTECT:
                    case BIFFRECORDTYPE.PROT4REVPASSWORD:
                        // IsProtected
                        break;
                    case BIFFRECORDTYPE.EOF:
                        _globals.SST?.ReadStrings();
                        return;

                    default:
                        continue;
                }
            }
        }

        private bool ReadWorkSheetGlobals(XlsWorksheet sheet)
        {
            XlsBiffIndex idx = null;

            _stream.Seek((int)sheet.DataOffset, SeekOrigin.Begin);

            XlsBiffBOF bof = _stream.Read() as XlsBiffBOF;
            if (bof == null || bof.Type != BIFFTYPE.Worksheet)
                return false;

            //// DumpBiffRecords();

            XlsBiffRecord rec = _stream.Read();
            if (rec == null || rec is XlsBiffEof)
                return false;

            if (rec is XlsBiffIndex)
            {
                idx = rec as XlsBiffIndex;
            }
            else if (rec is XlsBiffUncalced)
            {
                // Sometimes this come before the index...
                rec = _stream.Read();
                if (rec == null || rec is XlsBiffEof)
                    return false;

                idx = rec as XlsBiffIndex;
            }

            if (idx != null)
            {
                LogManager.Log(this).Debug("INDEX IsV8={0}", idx.IsV8);

                if (idx.LastExistingRow <= idx.FirstExistingRow)
                    return false;
            }

            while (!(rec is XlsBiffRow) && !(rec is XlsBiffBlankCell))
            {
                if (rec is XlsBiffDimensions dims)
                {
                    LogManager.Log(this).Debug("dims IsV8={0}", IsV8());
                    FieldCount = dims.LastColumn - 1;
                    break;
                }

                rec = _stream.Read();
            }

            // Handle when dimensions report less columns than used by cell records.
            int maxCellColumn = 0;
            while (rec != null && !(rec is XlsBiffEof))
            {
                if (rec is XlsBiffBlankCell cell)
                {
                    maxCellColumn = Math.Max(maxCellColumn, cell.ColumnIndex + 1);
                }

                rec = _stream.Read();
            }
            
            if (FieldCount < maxCellColumn)
                FieldCount = maxCellColumn;

            _stream.Seek((int)sheet.DataOffset, SeekOrigin.Begin);

            Depth = 0;

            return true;
        }

        /*private void DumpBiffRecords()
        {
            XlsBiffRecord rec = null;
            var startPos = stream.Position;

            do
            {
                rec = stream.Read();
                LogManager.Log(this).Debug(rec.Id.ToString());
            } while (rec != null && stream.Position < stream.Size);

            stream.Seek(startPos, SeekOrigin.Begin);
        }*/

        private void ReadNextBlock()
        {
            if (_reachedEndOfSheet)
                return;

            _currentRows.Clear();

            int currentRow = -1;

            object[] cellValues = null;

            XlsBiffRecord rec;

            while ((rec = _stream.Read()) != null)
            {
                if (rec is XlsBiffEof)
                {
                    _reachedEndOfSheet = true;
                    break;
                }

                if (rec is XlsBiffMSODrawing || rec is XlsBiffDbCell)
                {
                    break;
                }

                var cell = rec as XlsBiffBlankCell;
                if (cell != null)
                {
                    _largestObservedRow = Math.Max(_largestObservedRow, cell.RowIndex);

                    // In most cases cells are grouped by row
                    if (currentRow != cell.RowIndex)
                    {
                        if (!_currentRows.TryGetValue(cell.RowIndex, out cellValues))
                        {
                            cellValues = new object[FieldCount];
                            _currentRows.Add(cell.RowIndex, cellValues);
                        }

                        currentRow = cell.RowIndex;
                    }

                    PushCellValue(cellValues, cell);
                }
            }
        }

        /// <summary>
        /// Read a worksheet row.
        /// </summary>
        /// <returns>true if row was read successfully</returns>
        private bool ReadWorkSheetRow()
        {
            if (_currentRows.Count == 0 || Depth > _largestObservedRow)
            {
                ReadNextBlock();
            }

            if (Depth <= _largestObservedRow)
            {
                if (!_currentRows.TryGetValue(Depth, out _cellsValues))
                    _cellsValues = new object[FieldCount];
                Depth++;
                return true;
            }

            return false;
        }

        private void PushCellValue(object[] cellValues, XlsBiffBlankCell cell)
        {
            double doubleValue;
            LogManager.Log(this).Debug("PushCellValue {0}", cell.Id);
            switch (cell.Id)
            {
                case BIFFRECORDTYPE.BOOLERR:
                    if (cell.ReadByte(7) == 0)
                        cellValues[cell.ColumnIndex] = cell.ReadByte(6) != 0;
                    break;
                case BIFFRECORDTYPE.BOOLERR_OLD:
                    if (cell.ReadByte(8) == 0)
                        cellValues[cell.ColumnIndex] = cell.ReadByte(7) != 0;
                    break;
                case BIFFRECORDTYPE.INTEGER:
                case BIFFRECORDTYPE.INTEGER_OLD:
                    cellValues[cell.ColumnIndex] = ((XlsBiffIntegerCell)cell).Value;
                    break;
                case BIFFRECORDTYPE.NUMBER:
                case BIFFRECORDTYPE.NUMBER_OLD:

                    doubleValue = ((XlsBiffNumberCell)cell).Value;

                    cellValues[cell.ColumnIndex] = !ConvertOaDate ?
                        doubleValue : TryConvertOADateTime(doubleValue, cell.XFormat);

                    LogManager.Log(this).Debug("VALUE: {0}", doubleValue);
                    break;
                case BIFFRECORDTYPE.LABEL:
                case BIFFRECORDTYPE.LABEL_OLD:
                case BIFFRECORDTYPE.RSTRING:

                    cellValues[cell.ColumnIndex] = ((XlsBiffLabelCell)cell).Value;

                    LogManager.Log(this).Debug("VALUE: {0}", cellValues[cell.ColumnIndex]);
                    break;
                case BIFFRECORDTYPE.LABELSST:
                    string tmp = _globals.SST.GetString(((XlsBiffLabelSSTCell)cell).SSTIndex);
                    LogManager.Log(this).Debug("VALUE: {0}", tmp);
                    cellValues[cell.ColumnIndex] = tmp;
                    break;
                case BIFFRECORDTYPE.RK:

                    doubleValue = ((XlsBiffRKCell)cell).Value;

                    cellValues[cell.ColumnIndex] = !ConvertOaDate ?
                        doubleValue : TryConvertOADateTime(doubleValue, cell.XFormat);

                    LogManager.Log(this).Debug("VALUE: {0}", doubleValue);
                    break;
                case BIFFRECORDTYPE.MULRK:

                    XlsBiffMulRKCell rkCell = (XlsBiffMulRKCell)cell;
                    for (ushort j = cell.ColumnIndex; j <= rkCell.LastColumnIndex; j++)
                    {
                        doubleValue = rkCell.GetValue(j);
                        LogManager.Log(this).Debug("VALUE[{1}]: {0}", doubleValue, j);
                        cellValues[j] = !ConvertOaDate ? doubleValue : TryConvertOADateTime(doubleValue, rkCell.GetXF(j));
                    }

                    break;
                case BIFFRECORDTYPE.BLANK:
                case BIFFRECORDTYPE.BLANK_OLD:
                case BIFFRECORDTYPE.MULBLANK:
                    // Skip blank cells
                    break;
                case BIFFRECORDTYPE.FORMULA:
                case BIFFRECORDTYPE.FORMULA_OLD:

                    object objectValue = ((XlsBiffFormulaCell)cell).Value;

                    if (objectValue is FORMULAERROR)
                    {
                        objectValue = null;
                    }
                    else
                    {
                        cellValues[cell.ColumnIndex] = !ConvertOaDate ?
                            objectValue : TryConvertOADateTime(objectValue, cell.XFormat); // date time offset
                    }

                    LogManager.Log(this).Debug("VALUE: {0}", objectValue);
                    break;
            }
        }

        private bool InitializeSheetRead()
        {
            if (_sheetIndex == -1)
                _sheetIndex = 0;

            if (!ReadWorkSheetGlobals(_sheets[_sheetIndex]))
            {
                return false;
            }

            return true;
        }

        private object TryConvertOADateTime(double value, ushort xFormat)
        {
            ushort format;
            if (xFormat < _globals.ExtendedFormats.Count)
            {
                // If a cell XF record does not contain explicit attributes in a group (if the attribute group flag is not set),
                // it repeats the attributes of its style XF record.
                var rec = _globals.ExtendedFormats[xFormat];
                switch (rec.Id)
                {
                    case BIFFRECORDTYPE.XF_V2:
                        format = (ushort)(rec.ReadByte(2) & 0x3F);
                        break;
                    case BIFFRECORDTYPE.XF_V3:
                        format = rec.ReadByte(1);
                        break;
                    case BIFFRECORDTYPE.XF_V4:
                        format = rec.ReadByte(1);
                        break;

                    default:
                        format = rec.ReadUInt16(2);
                        break;
                }
            }
            else
            {
                format = xFormat;
            }

            switch (format)
            {
                // numeric built in formats
                case 0: // "General";
                case 1: // "0";
                case 2: // "0.00";
                case 3: // "#,##0";
                case 4: // "#,##0.00";
                case 5: // "\"$\"#,##0_);(\"$\"#,##0)";
                case 6: // "\"$\"#,##0_);[Red](\"$\"#,##0)";
                case 7: // "\"$\"#,##0.00_);(\"$\"#,##0.00)";
                case 8: // "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
                case 9: // "0%";
                case 10: // "0.00%";
                case 11: // "0.00E+00";
                case 12: // "# ?/?";
                case 13: // "# ??/??";
                case 0x30: // "##0.0E+0";

                case 0x25: // "_(#,##0_);(#,##0)";
                case 0x26: // "_(#,##0_);[Red](#,##0)";
                case 0x27: // "_(#,##0.00_);(#,##0.00)";
                case 40: // "_(#,##0.00_);[Red](#,##0.00)";
                case 0x29: // "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)";
                case 0x2a: // "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)";
                case 0x2b: // "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)";
                case 0x2c: // "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)";
                    return value;

                // date formats
                case 14: // this.GetDefaultDateFormat();
                case 15: // "D-MM-YY";
                case 0x10: // "D-MMM";
                case 0x11: // "MMM-YY";
                case 0x12: // "h:mm AM/PM";
                case 0x13: // "h:mm:ss AM/PM";
                case 20: // "h:mm";
                case 0x15: // "h:mm:ss";
                case 0x16: // string.Format("{0} {1}", this.GetDefaultDateFormat(), this.GetDefaultTimeFormat());

                case 0x2d: // "mm:ss";
                case 0x2e: // "[h]:mm:ss";
                case 0x2f: // "mm:ss.0";
                    return Helpers.ConvertFromOATime(value);
                case 0x31: // "@";
                    return value.ToString(CultureInfo.CurrentCulture); // TODO: What is the exepcted culture here?

                default:
                    if (_globals.Formats.TryGetValue(format, out XlsBiffFormatString fmtString))
                    {
                        var fmt = fmtString.Value;
                        var formatReader = new FormatReader { FormatString = fmt };
                        if (formatReader.IsDateFormatString())
                            return Helpers.ConvertFromOATime(value);
                    }

                    return value;
            }
        }

        private object TryConvertOADateTime(object value, ushort xFormat)
        {
            double doubleValue;

            if (value == null)
                return null;

            if (double.TryParse(value.ToString(), out doubleValue))
                return TryConvertOADateTime(doubleValue, xFormat);

            return value;
        }

        private void ResetSheetData()
        {
            FieldCount = 0;
            _isFirstRead = true;
            _reachedEndOfSheet = false;
            _currentRows.Clear();
            _largestObservedRow = -1;
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

        public Encoding Encoding { get; private set; }

        public string Name
        {
            get
            {
                if (_sheets != null && _sheets.Count > 0)
                    return _sheets[_sheetIndex].Name;

                return null;
            }
        }

        public string VisibleState
        {
            get
            {
                if (_sheets != null && _sheets.Count > 0)
                    return _sheets[_sheetIndex].VisibleState;

                return null;
            }
        }

        public int Depth { get; private set; }

        public int ResultsCount => _globals.Sheets.Count;

        public bool IsClosed { get; private set; }

        public int FieldCount { get; private set; }

        public int RecordsAffected => throw new NotSupportedException();
        
        public object this[int i] => _cellsValues[i];

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
            return _cellsValues[i] == null ? null : _cellsValues[i].GetType();
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
            _sheetIndex = 0;

            ResetSheetData();
        }

        public void Close()
        {
            if (IsClosed)
                return;

            if (_sheets != null)
            {
                _sheets.Clear();
                _sheets = null;
            }

            // m_workbookData = null;
            _stream = null;
            _globals = null;
            Encoding = null;
            _hdr = null;

            if (_file != null)
            {
                _file.Dispose();
                _file = null;
            }

            IsClosed = true;
        }

        public bool NextResult()
        {
            if (_sheetIndex >= ResultsCount - 1)
                return false;

            _sheetIndex++;

            ResetSheetData();

            return true;
        }

        public bool Read()
        {
            if (_isFirstRead)
            {
                _isFirstRead = false;
                if (!InitializeSheetRead())
                    return false;

                if (IsFirstRowAsColumnNames)
                {
                    if (ReadWorkSheetRow())
                    {
                        _cellsNames = new string[_cellsValues.Length];
                        for (var i = 0; i < _cellsValues.Length; i++)
                        {
                            var value = _cellsValues[i]?.ToString();
                            if (!string.IsNullOrEmpty(value))
                            {
                                _cellsNames[i] = value;
                            }
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    _cellsNames = null; // no columns
                }
            }

            return ReadWorkSheetRow();
        }

        public bool GetBoolean(int i)
        {
            return !IsDBNull(i) && bool.Parse(_cellsValues[i].ToString());
        }

        public DateTime GetDateTime(int i)
        {
            if (IsDBNull(i))
                return DateTime.MinValue;

            // requested change: 3
            object val = _cellsValues[i];

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
    }
}