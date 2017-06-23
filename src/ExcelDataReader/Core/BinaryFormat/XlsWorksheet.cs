using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader.Log;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Worksheet section in workbook
    /// </summary>
    internal class XlsWorksheet : IWorksheet
    {
        public XlsWorksheet(XlsWorkbook workbook, int index)
        {
            Workbook = workbook;
            Index = index;

            var refSheet = workbook.Sheets[index];
            Name = refSheet.SheetName;
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

        /// <summary>
        /// Gets the visibility of worksheet
        /// </summary>
        public string VisibleState { get; }

        /// <summary>
        /// Gets the zero-based index of worksheet
        /// </summary>
        public int Index { get; }

        /// <summary>
        /// Gets the worksheet data offset.
        /// </summary>
        public uint DataOffset { get; }
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

        public XlsWorkbook Workbook { get; }

        public XlsBiffStream BiffStream => Workbook.BiffStream;

        public IEnumerable<object[]> ReadRows()
        {
            var rowIndex = 0;
            BiffStream.Seek((int)DataOffset, SeekOrigin.Begin);

            while (true)
            {
                var block = ReadNextBlock();

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

        private XlsRowBlock ReadNextBlock()
        {
            var result = new XlsRowBlock();
            result.Rows = new Dictionary<int, object[]>();

            var currentRowIndex = -1;
            var currentRow = (object[])null;

            XlsBiffRecord rec;

            while ((rec = BiffStream.Read()) != null)
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

                var cell = rec as XlsBiffBlankCell;
                if (cell != null)
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

                    var additionalRecords = new List<XlsBiffRecord>();
                    while (!PushCellValue(currentRow, cell, additionalRecords))
                    {
                        var additionalRecord = BiffStream.Read();
                        additionalRecords.Add(additionalRecord);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Returns false if more records are needed to parse the value. The caller is expected to retry after parsing a record into additionalRecords.
        /// </summary>
        private bool PushCellValue(object[] cellValues, XlsBiffBlankCell cell, List<XlsBiffRecord> additionalRecords)
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

                    cellValues[cell.ColumnIndex] = !Workbook.ConvertOaDate ?
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
                    string tmp = Workbook.SST.GetString(((XlsBiffLabelSSTCell)cell).SSTIndex);
                    LogManager.Log(this).Debug("VALUE: {0}", tmp);
                    cellValues[cell.ColumnIndex] = tmp;
                    break;
                case BIFFRECORDTYPE.RK:

                    doubleValue = ((XlsBiffRKCell)cell).Value;

                    cellValues[cell.ColumnIndex] = !Workbook.ConvertOaDate ?
                        doubleValue : TryConvertOADateTime(doubleValue, cell.XFormat);

                    LogManager.Log(this).Debug("VALUE: {0}", doubleValue);
                    break;
                case BIFFRECORDTYPE.MULRK:

                    XlsBiffMulRKCell rkCell = (XlsBiffMulRKCell)cell;
                    for (ushort j = cell.ColumnIndex; j <= rkCell.LastColumnIndex; j++)
                    {
                        doubleValue = rkCell.GetValue(j);
                        LogManager.Log(this).Debug("VALUE[{1}]: {0}", doubleValue, j);
                        cellValues[j] = !Workbook.ConvertOaDate ? doubleValue : TryConvertOADateTime(doubleValue, rkCell.GetXF(j));
                    }

                    break;
                case BIFFRECORDTYPE.BLANK:
                case BIFFRECORDTYPE.BLANK_OLD:
                case BIFFRECORDTYPE.MULBLANK:
                    // Skip blank cells
                    break;
                case BIFFRECORDTYPE.FORMULA:
                case BIFFRECORDTYPE.FORMULA_OLD:

                    var objectValue = (object)null;
                    if (!TryGetFormulaValue((XlsBiffFormulaCell)cell, additionalRecords, out objectValue))
                    {
                        // want additional records
                        return false;
                    }

                    cellValues[cell.ColumnIndex] = !Workbook.ConvertOaDate ?
                            objectValue : TryConvertOADateTime(objectValue, cell.XFormat); // date time offset;
                    LogManager.Log(this).Debug("VALUE: {0}", objectValue);
                    break;
            }

            return true;
        }

        private bool TryGetFormulaValue(XlsBiffFormulaCell formulaCell, List<XlsBiffRecord> additionalRecords, out object result)
        {
            if (formulaCell.IsBoolean)
            {
                result = formulaCell.BooleanValue;
                return true;
            }

            if (formulaCell.IsError)
            {
                result = null;
                return true;
            }
            else if (formulaCell.IsEmptyString)
            {
                result = string.Empty;
                return true;
            }
            else if (formulaCell.IsXNum)
            {
                result = formulaCell.XNumValue;
                return true;
            }
            else if (formulaCell.IsString)
            {
                if (additionalRecords.Count == 0)
                {
                    result = null;
                    return false;
                }

                if (additionalRecords.Count == 1)
                {
                    var recId = additionalRecords[0].Id;
                    if (recId == BIFFRECORDTYPE.SHAREDFMLA)
                    {
                        result = null;
                        return false;
                    }
                    else if (recId == BIFFRECORDTYPE.STRING)
                    {
                        var stringRecord = additionalRecords[0] as XlsBiffFormulaString;
                        result = stringRecord.Value;
                        return true;
                    }
                }

                {
                    var recId = additionalRecords[1].Id;
                    if (recId == BIFFRECORDTYPE.STRING)
                    {
                        var stringRecord = additionalRecords[1] as XlsBiffFormulaString;
                        result = stringRecord.Value;
                        return true;
                    }
                }

                // Bad data - could not find a string following the formula
                result = null;
                return true;
            }
            else
            {
                // Bad data or new formula value type
                result = null;
                return true;
            }
        }

        private object TryConvertOADateTime(double value, ushort xFormat)
        {
            ushort format;
            if (xFormat < Workbook.ExtendedFormats.Count)
            {
                // If a cell XF record does not contain explicit attributes in a group (if the attribute group flag is not set),
                // it repeats the attributes of its style XF record.
                var rec = Workbook.ExtendedFormats[xFormat];
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
                    return value.ToString(); // TODO: What is the exepcted culture here?

                default:
                    if (Workbook.Formats.TryGetValue(format, out XlsBiffFormatString fmtString))
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

        private object ReadWorksheetGlobals()
        {
            XlsBiffIndex idx = null;

            BiffStream.Seek((int)DataOffset, SeekOrigin.Begin);

            XlsBiffBOF bof = BiffStream.Read() as XlsBiffBOF;
            if (bof == null || bof.Type != BIFFTYPE.Worksheet)
                return null;

            //// DumpBiffRecords();

            XlsBiffRecord rec = BiffStream.Read();
            if (rec == null || rec is XlsBiffEof)
                return null;

            if (rec is XlsBiffIndex)
            {
                idx = rec as XlsBiffIndex;
            }
            else if (rec is XlsBiffUncalced)
            {
                // Sometimes this come before the index...
                rec = BiffStream.Read();
                if (rec == null || rec is XlsBiffEof)
                    return null;

                idx = rec as XlsBiffIndex;
            }

            if (idx != null)
            {
                LogManager.Log(this).Debug("INDEX IsV8={0}", idx.IsV8);

                if (idx.LastExistingRow <= idx.FirstExistingRow)
                    return null;
            }

            while (!(rec is XlsBiffRow) && !(rec is XlsBiffBlankCell))
            {
                if (rec is XlsBiffDimensions dims)
                {
                    // LogManager.Log(this).Debug("dims IsV8={0}", IsV8());
                    FieldCount = dims.LastColumn - 1;
                    break;
                }

                rec = BiffStream.Read();
            }

            // Handle when dimensions report less columns than used by cell records.
            int maxCellColumn = 0;
            Dictionary<int, bool> previousBlocksObservedRows = new Dictionary<int, bool>();
            Dictionary<int, bool> observedRows = new Dictionary<int, bool>();
            while (rec != null && !(rec is XlsBiffEof))
            {
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

                rec = BiffStream.Read();
            }

            if (FieldCount < maxCellColumn)
                FieldCount = maxCellColumn;

            return true;
        }

        internal class XlsRowBlock
        {
            public Dictionary<int, object[]> Rows { get; set; }

            public bool EndOfSheet { get; set; }
        }
    }
}