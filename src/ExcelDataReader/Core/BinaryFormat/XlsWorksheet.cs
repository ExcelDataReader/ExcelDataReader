using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Log;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Worksheet section in workbook
    /// </summary>
    internal class XlsWorksheet : IWorksheet
    {
        public XlsWorksheet(XlsWorkbook workbook, XlsBiffBoundSheet refSheet, Stream stream)
        {
            Workbook = workbook;
            Stream = stream;

            IsDate1904 = workbook.IsDate1904;
            Formats = new Dictionary<ushort, XlsBiffFormatString>(workbook.Formats);
            ExtendedFormats = new List<XlsBiffRecord>(workbook.ExtendedFormats);
            Encoding = workbook.Encoding;

            Name = refSheet.GetSheetName(workbook.Encoding);
            DataOffset = refSheet.StartOffset;

            switch (refSheet.VisibleState)
            {
                case XlsBiffBoundSheet.SheetVisibility.Hidden:
                    VisibleState = "hidden";
                    break;
                case XlsBiffBoundSheet.SheetVisibility.VeryHidden:
                    VisibleState = "veryhidden";
                    break;
                default:
                    VisibleState = "visible";
                    break;
            }

            ReadWorksheetGlobals();
        }

        /// <summary>
        /// Gets the worksheet name
        /// </summary>
        public string Name { get; }

        public string CodeName { get; private set; }

        /// <summary>
        /// Gets the visibility of worksheet
        /// </summary>
        public string VisibleState { get; }

        public HeaderFooter HeaderFooter { get; private set; }

        /// <summary>
        /// Gets the worksheet data offset.
        /// </summary>
        public uint DataOffset { get; }

        public Stream Stream { get; }

        public Dictionary<ushort, XlsBiffFormatString> Formats { get; }

        public List<XlsBiffRecord> ExtendedFormats { get; }

        public Encoding Encoding { get; private set; }

/*
    TODO: populate these in ReadWorksheetGlobals() if needed
        public XlsBiffSimpleValueRecord CalcMode { get; set; }

        public XlsBiffSimpleValueRecord CalcCount { get; set; }

        public XlsBiffSimpleValueRecord RefMode { get; set; }

        public XlsBiffSimpleValueRecord Iteration { get; set; }

        public XlsBiffRecord Delta { get; set; }
        
        public XlsBiffRecord Window { get; set; }
*/

        public int FieldCount { get; private set; }

        public bool RowContentInMultipleBlocks { get; private set; }

        public bool IsDate1904 { get; private set; }

        public XlsWorkbook Workbook { get; }

        public IEnumerable<object[]> ReadRows()
        {
            var rowIndex = 0;
            var biffStream = new XlsBiffStream(Stream, (int)DataOffset, Workbook.BiffVersion, null, Workbook.SecretKey, Workbook.Encryption);

            while (true)
            {
                var block = ReadNextBlock(biffStream);

                var maxRow = int.MinValue;
                foreach (var blockRowIndex in block.Rows.Keys)
                {
                    maxRow = Math.Max(maxRow, blockRowIndex);
                }

                for (; rowIndex <= maxRow; rowIndex++)
                {
                    if (block.Rows.TryGetValue(rowIndex, out var row))
                    {
                        yield return row;
                    }
                    else
                    {
                        row = new object[FieldCount];
                        yield return row;
                    }
                }

                if (block.EndOfSheet || block.Rows.Count == 0)
                {
                    break;
                }
            }
        }

        private XlsRowBlock ReadNextBlock(XlsBiffStream biffStream)
        {
            var result = new XlsRowBlock { Rows = new Dictionary<int, object[]>() };

            var currentRowIndex = -1;
            object[] currentRow = null;

            XlsBiffRecord rec;
            XlsBiffRecord ixfe = null;

            while ((rec = biffStream.Read()) != null)
            {
                if (rec is XlsBiffEof)
                {
                    result.EndOfSheet = true;
                    break;
                }

                if (rec is XlsBiffMSODrawing || (!RowContentInMultipleBlocks && rec is XlsBiffDbCell))
                {
                    break;
                }

                if (rec.Id == BIFFRECORDTYPE.IXFE)
                {
                    // BIFF2: If cell.xformat == 63, this contains the actual XF index >= 63
                    ixfe = rec;
                }

                if (rec is XlsBiffBlankCell cell)
                {
                    // In most cases cells are grouped by row
                    if (currentRowIndex != cell.RowIndex)
                    {
                        if (!result.Rows.TryGetValue(cell.RowIndex, out currentRow))
                        {
                            currentRow = new object[FieldCount];
                            result.Rows.Add(cell.RowIndex, currentRow);
                        }

                        currentRowIndex = cell.RowIndex;
                    }

                    ushort xFormat;
                    if (Workbook.BiffVersion == 2 && cell.XFormat == 63 && ixfe != null)
                    {
                        xFormat = ixfe.ReadUInt16(0);
                    }
                    else
                    {
                        xFormat = cell.XFormat;
                    }

                    var additionalRecords = new List<XlsBiffRecord>();
                    while (!PushCellValue(currentRow, cell, xFormat, additionalRecords))
                    {
                        var additionalRecord = biffStream.Read();
                        additionalRecords.Add(additionalRecord);
                    }

                    ixfe = null;
                }
            }

            return result;
        }

        /// <summary>
        /// Returns false if more records are needed to parse the value. The caller is expected to retry after parsing a record into additionalRecords.
        /// </summary>
        private bool PushCellValue(object[] cellValues, XlsBiffBlankCell cell, ushort xFormat, List<XlsBiffRecord> additionalRecords)
        {
            double doubleValue;
            int intValue;
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
                    intValue = ((XlsBiffIntegerCell)cell).Value;
                    cellValues[cell.ColumnIndex] = TryConvertOADateTime(intValue, xFormat);
                    break;
                case BIFFRECORDTYPE.NUMBER:
                case BIFFRECORDTYPE.NUMBER_OLD:
                    doubleValue = ((XlsBiffNumberCell)cell).Value;
                    cellValues[cell.ColumnIndex] = TryConvertOADateTime(doubleValue, xFormat);
                    LogManager.Log(this).Debug("VALUE: {0}", doubleValue);
                    break;
                case BIFFRECORDTYPE.LABEL:
                case BIFFRECORDTYPE.LABEL_OLD:
                case BIFFRECORDTYPE.RSTRING:
                    cellValues[cell.ColumnIndex] = ((XlsBiffLabelCell)cell).GetValue(Encoding);
                    LogManager.Log(this).Debug("VALUE: {0}", cellValues[cell.ColumnIndex]);
                    break;
                case BIFFRECORDTYPE.LABELSST:
                    string tmp = Workbook.SST.GetString(((XlsBiffLabelSSTCell)cell).SSTIndex, Encoding);
                    LogManager.Log(this).Debug("VALUE: {0}", tmp);
                    cellValues[cell.ColumnIndex] = tmp;
                    break;
                case BIFFRECORDTYPE.RK:
                    doubleValue = ((XlsBiffRKCell)cell).Value;
                    cellValues[cell.ColumnIndex] = TryConvertOADateTime(doubleValue, xFormat);
                    LogManager.Log(this).Debug("VALUE: {0}", doubleValue);
                    break;
                case BIFFRECORDTYPE.MULRK:

                    XlsBiffMulRKCell rkCell = (XlsBiffMulRKCell)cell;
                    ushort lastColumnIndex = rkCell.LastColumnIndex;
                    for (ushort j = cell.ColumnIndex; j <= lastColumnIndex; j++)
                    {
                        doubleValue = rkCell.GetValue(j);
                        LogManager.Log(this).Debug("VALUE[{1}]: {0}", doubleValue, j);
                        cellValues[j] = TryConvertOADateTime(doubleValue, rkCell.GetXF(j));
                    }

                    break;
                case BIFFRECORDTYPE.BLANK:
                case BIFFRECORDTYPE.BLANK_OLD:
                case BIFFRECORDTYPE.MULBLANK:
                    // Skip blank cells
                    break;
                case BIFFRECORDTYPE.FORMULA:
                case BIFFRECORDTYPE.FORMULA_V3:
                case BIFFRECORDTYPE.FORMULA_V4:
                    if (!TryGetFormulaValue((XlsBiffFormulaCell)cell, xFormat, additionalRecords, out object objectValue))
                    {
                        // want additional records
                        return false;
                    }

                    cellValues[cell.ColumnIndex] = objectValue;
                    LogManager.Log(this).Debug("VALUE: {0}", objectValue);
                    break;
            }

            return true;
        }

        private bool TryGetFormulaValue(XlsBiffFormulaCell formulaCell, ushort xFormat, List<XlsBiffRecord> additionalRecords, out object result)
        {
            switch (formulaCell.FormulaType)
            {
                case XlsBiffFormulaCell.FormulaValueType.Boolean:
                    result = formulaCell.BooleanValue;
                    return true;
                case XlsBiffFormulaCell.FormulaValueType.Error:
                    result = null;
                    return true;
                case XlsBiffFormulaCell.FormulaValueType.EmptyString:
                    result = string.Empty;
                    return true;
                case XlsBiffFormulaCell.FormulaValueType.Number:
                    result = TryConvertOADateTime(formulaCell.XNumValue, xFormat);
                    return true;
                case XlsBiffFormulaCell.FormulaValueType.String when additionalRecords.Count == 0:
                    result = null;

                    // Request additional records.
                    return false;
                case XlsBiffFormulaCell.FormulaValueType.String:
                    BIFFRECORDTYPE recId;

                    if (additionalRecords.Count == 1)
                    {
                        recId = additionalRecords[0].Id;
                        if (recId == BIFFRECORDTYPE.SHAREDFMLA)
                        {
                            result = null;

                            // Request additional records.
                            return false;
                        }

                        if (recId == BIFFRECORDTYPE.STRING)
                        {
                            var stringRecord = (XlsBiffFormulaString)additionalRecords[0];
                            result = stringRecord.GetValue(Encoding);
                            return true;
                        }
                    }

                    // The old implementation would throw an IndexOutOfRangeException if the record isn't
                    // a SHAREDFMLA or STRING. 
                    if (additionalRecords.Count > 1)
                    {
                        recId = additionalRecords[1].Id;
                        if (recId == BIFFRECORDTYPE.STRING)
                        {
                            var stringRecord = (XlsBiffFormulaString)additionalRecords[1];
                            result = stringRecord.GetValue(Encoding);
                            return true;
                        }
                    }

                    // Bad data - could not find a string following the formula
                    break;
            }

            // Bad data or new formula value type
            result = null;
            return true;
        }

        private object TryConvertOADateTime(double value, ushort xFormat)
        {
            if (IsDateFormat(xFormat))
                return Helpers.ConvertFromOATime(value, IsDate1904);
            return value;
        }

        private object TryConvertOADateTime(int value, ushort xFormat)
        {
            if (IsDateFormat(xFormat))
                return Helpers.ConvertFromOATime(value, IsDate1904);
            return value;
        }

        private bool IsDateFormat(ushort xFormat)
        {
            ushort format;
            if (xFormat < ExtendedFormats.Count)
            {
                // If a cell XF record does not contain explicit attributes in a group (if the attribute group flag is not set),
                // it repeats the attributes of its style XF record.
                var rec = ExtendedFormats[xFormat];
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

            // From BIFF5 on, the built-in number formats will be omitted. 
            if (Workbook.BiffVersion >= 5)
            {
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
                        return false;

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
                        return true;
                    case 0x31: // "@";
                        return false; // NOTE: was value.ToString();
                }
            }

            if (Formats.TryGetValue(format, out XlsBiffFormatString fmtString))
            {
                var fmt = fmtString.GetValue(Encoding);
                var formatReader = new FormatReader { FormatString = fmt };
                return formatReader.IsDateFormatString();
            }

            return false;
        }

        private void ReadWorksheetGlobals()
        {
            var biffStream = new XlsBiffStream(Stream, (int)DataOffset, Workbook.BiffVersion, null, Workbook.SecretKey, Workbook.Encryption);
            
            // Check the expected BOF record was found in the BIFF stream
            if (biffStream.BiffVersion == 0 || biffStream.BiffType != BIFFTYPE.Worksheet)
                return;

            XlsBiffHeaderFooterString header = null;
            XlsBiffHeaderFooterString footer = null;

            // Handle when dimensions report less columns than used by cell records.
            int maxCellColumn = 0;
            Dictionary<int, bool> previousBlocksObservedRows = new Dictionary<int, bool>();
            Dictionary<int, bool> observedRows = new Dictionary<int, bool>();

            XlsBiffRecord rec = biffStream.Read();
            while (rec != null && !(rec is XlsBiffEof))
            {
                if (rec is XlsBiffDimensions dims)
                {
                    FieldCount = dims.LastColumn;
                }

                if (rec.Id == BIFFRECORDTYPE.RECORD1904)
                {
                    IsDate1904 = ((XlsBiffSimpleValueRecord)rec).Value == 1;
                }

                if (rec.Id == BIFFRECORDTYPE.XF_V2 || rec.Id == BIFFRECORDTYPE.XF_V3 || rec.Id == BIFFRECORDTYPE.XF_V4)
                {
                    ExtendedFormats.Add(rec);
                }

                if (rec.Id == BIFFRECORDTYPE.FORMAT)
                {
                    var fmt = (XlsBiffFormatString)rec;
                    if (Workbook.BiffVersion >= 5)
                    {
                        // fmt.Index exists on BIFF5+ only
                        Formats.Add(fmt.Index, fmt);
                    }
                    else
                    {
                        Formats.Add((ushort)Formats.Count, fmt);
                    }
                }

                if (rec.Id == BIFFRECORDTYPE.FORMAT_V23)
                {
                    var fmt = (XlsBiffFormatString)rec;
                    Formats.Add((ushort)Formats.Count, fmt);
                }

                if (rec.Id == BIFFRECORDTYPE.CODEPAGE)
                {
                    var codePage = (XlsBiffSimpleValueRecord)rec;
                    Encoding = EncodingHelper.GetEncoding(codePage.Value);
                }

                if (rec.Id == BIFFRECORDTYPE.HEADER && rec.RecordSize > 0)
                {
                    header = (XlsBiffHeaderFooterString)rec;
                }

                if (rec.Id == BIFFRECORDTYPE.FOOTER && rec.RecordSize > 0)
                {
                    footer = (XlsBiffHeaderFooterString)rec;
                }

                if (rec.Id == BIFFRECORDTYPE.CODENAME)
                {
                    var codeName = (XlsBiffCodeName)rec;
                    CodeName = codeName.GetValue(Encoding);
                }

                if (!RowContentInMultipleBlocks && rec is XlsBiffDbCell)
                {
                    foreach (int row in observedRows.Keys)
                    {
                        previousBlocksObservedRows[row] = true;
                    }
                    
                    observedRows.Clear();
                }

                if (rec is XlsBiffBlankCell cell)
                {
                    maxCellColumn = Math.Max(maxCellColumn, cell.ColumnIndex + 1);

                    if (!RowContentInMultipleBlocks)
                    {
                        if (previousBlocksObservedRows.ContainsKey(cell.RowIndex))
                        {
                            RowContentInMultipleBlocks = true;
                            previousBlocksObservedRows.Clear();
                            observedRows.Clear();
                        }

                        observedRows[cell.RowIndex] = true;
                    }
                }

                rec = biffStream.Read();
            }

            if (header != null || footer != null)
            {
                HeaderFooter = new HeaderFooter(false, false)
                {
                    OddHeader = header?.GetValue(Encoding),
                    OddFooter = footer?.GetValue(Encoding),
                };
            }

            if (FieldCount < maxCellColumn)
                FieldCount = maxCellColumn;
        }

        internal class XlsRowBlock
        {
            public Dictionary<int, object[]> Rows { get; set; }

            public bool EndOfSheet { get; set; }
        }
    }
}