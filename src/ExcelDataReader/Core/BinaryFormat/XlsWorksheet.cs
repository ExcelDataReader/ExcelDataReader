using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Core.NumberFormat;
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
            Formats = new Dictionary<ushort, NumberFormatString>(workbook.Formats);
            ExtendedFormats = new List<XlsBiffXF>(workbook.ExtendedFormats);
            Encoding = workbook.Encoding;
            RowMinMaxOffsets = new Dictionary<int, KeyValuePair<int, int>>();
            DefaultRowHeight = 255; // 12.75 points

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

        public CellRange[] MergeCells { get; private set; }

        /// <summary>
        /// Gets the worksheet data offset.
        /// </summary>
        public uint DataOffset { get; }

        public Stream Stream { get; }

        public Dictionary<ushort, NumberFormatString> Formats { get; }

        public List<XlsBiffXF> ExtendedFormats { get; }

        public Encoding Encoding { get; private set; }

        public double DefaultRowHeight { get; set; }

        public Dictionary<int, KeyValuePair<int, int>> RowMinMaxOffsets { get; }

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

        public int RowCount { get; private set; }

        public bool IsDate1904 { get; private set; }

        public XlsWorkbook Workbook { get; }
                
        public IEnumerable<Row> ReadRows()
        {
            var rowIndex = 0;
            using (var biffStream = new XlsBiffStream(Stream, (int)DataOffset, Workbook.BiffVersion, null, Workbook.SecretKey, Workbook.Encryption))
            {
                foreach (var rowBlock in ReadWorksheetRows(biffStream))
                {
                    for (; rowIndex < rowBlock.RowIndex; ++rowIndex)
                    {
                        yield return new Row()
                        {
                            RowIndex = rowIndex,
                            Height = DefaultRowHeight / 20.0,
                            Cells = new List<Cell>()
                        };
                    }

                    rowIndex++;
                    yield return rowBlock;
                }
            }
        }

        public NumberFormatString GetNumberFormatString(int numberFormatIndex)
        {
            if (Formats.TryGetValue((ushort)numberFormatIndex, out var fmtString))
            {
                return fmtString;
            }
            else
            {
                return BuiltinNumberFormat.GetBuiltinNumberFormat(numberFormatIndex);
            }
        }

        private IEnumerable<Row> ReadWorksheetRows(XlsBiffStream biffStream)
        {
            var rowIndex = 0;

            while (rowIndex < RowCount)
            {
                // Read up to 32 rows at a time
                var blockRowCount = Math.Min(32, RowCount - rowIndex);

                var block = ReadNextBlock(biffStream, rowIndex, blockRowCount);
                
                for (var i = 0; i < blockRowCount; ++i)
                {
                    if (block.Rows.TryGetValue(rowIndex + i, out var row))
                    {
                        yield return row;
                    }
                }

                rowIndex += blockRowCount;
            }
        }

        private XlsRowBlock ReadNextBlock(XlsBiffStream biffStream, int startRow, int rows)
        {
            var result = new XlsRowBlock { Rows = new Dictionary<int, Row>() };

            XlsBiffRecord rec;
            XlsBiffRecord ixfe = null;

            if (!GetMinMaxOffsetsForRowBlock(startRow, rows, out var minOffset, out var maxOffset))
                return result;

            biffStream.Position = minOffset;

            while (biffStream.Position <= maxOffset && (rec = biffStream.Read()) != null)
            {
                if (rec.Id == BIFFRECORDTYPE.ROW || rec.Id == BIFFRECORDTYPE.ROW_V2)
                {
                    var rowRecord = (XlsBiffRow)rec;
                    var currentRow = EnsureRow(result, rowRecord.RowIndex);
                    currentRow.Height = (rowRecord.UseDefaultRowHeight ? DefaultRowHeight : rowRecord.RowHeight) / 20.0;
                }

                if (rec.Id == BIFFRECORDTYPE.IXFE)
                {
                    // BIFF2: If cell.xformat == 63, this contains the actual XF index >= 63
                    ixfe = rec;
                }

                if (rec.IsCell)
                {
                    var cell = (XlsBiffBlankCell)rec;
                    var currentRow = EnsureRow(result, cell.RowIndex);

                    if (cell.Id == BIFFRECORDTYPE.MULRK)
                    {
                        var cellValues = ReadMultiCell(cell);
                        currentRow.Cells.AddRange(cellValues);
                    }
                    else
                    {
                        ushort xFormat;
                        if (Workbook.BiffVersion == 2 && cell.XFormat == 63 && ixfe != null)
                        {
                            xFormat = ixfe.ReadUInt16(0);
                        }
                        else
                        {
                            xFormat = cell.XFormat;
                        }

                        var numberFormatIndex = GetNumberFormatIndex(xFormat);
                        var cellValue = ReadSingleCell(biffStream, cell, numberFormatIndex);
                        currentRow.Cells.Add(cellValue);
                    }

                    ixfe = null;
                }
            }
            
            return result;
        }
        
        private Row EnsureRow(XlsRowBlock result, int rowIndex)
        {
            if (!result.Rows.TryGetValue(rowIndex, out var currentRow))
            {
                currentRow = new Row()
                {
                    RowIndex = rowIndex,
                    Height = DefaultRowHeight / 20.0,
                    Cells = new List<Cell>()
                };

                result.Rows.Add(rowIndex, currentRow);
            }

            return currentRow;
        }

        private List<Cell> ReadMultiCell(XlsBiffBlankCell cell)
        {
            LogManager.Log(this).Debug("ReadMultiCell {0}", cell.Id);

            var result = new List<Cell>();
            switch (cell.Id)
            {
                case BIFFRECORDTYPE.MULRK:

                    XlsBiffMulRKCell rkCell = (XlsBiffMulRKCell)cell;
                    ushort lastColumnIndex = rkCell.LastColumnIndex;
                    for (ushort j = cell.ColumnIndex; j <= lastColumnIndex; j++)
                    {
                        var numberFormatIndex = GetNumberFormatIndex(rkCell.GetXF(j));
                        var resultCell = new Cell()
                        {
                            ColumnIndex = j,
                            Value = TryConvertOADateTime(rkCell.GetValue(j), numberFormatIndex),
                            NumberFormatIndex = numberFormatIndex
                        };

                        result.Add(resultCell);

                        LogManager.Log(this).Debug("VALUE[{1}]: {0}", resultCell.Value, j);
                    }

                    break;
            }

            return result;
        }

        /// <summary>
        /// Reads additional records if needed: a string record might follow a formula result
        /// </summary>
        private Cell ReadSingleCell(XlsBiffStream biffStream, XlsBiffBlankCell cell, int numberFormatIndex)
        {
            LogManager.Log(this).Debug("ReadSingleCell {0}", cell.Id);

            double doubleValue;
            int intValue;
            object objectValue;

            var result = new Cell()
            {
                ColumnIndex = cell.ColumnIndex,
                NumberFormatIndex = numberFormatIndex
            };

            switch (cell.Id)
            {
                case BIFFRECORDTYPE.BOOLERR:
                    if (cell.ReadByte(7) == 0)
                        result.Value = cell.ReadByte(6) != 0;
                    break;
                case BIFFRECORDTYPE.BOOLERR_OLD:
                    if (cell.ReadByte(8) == 0)
                        result.Value = cell.ReadByte(7) != 0;
                    break;
                case BIFFRECORDTYPE.INTEGER:
                case BIFFRECORDTYPE.INTEGER_OLD:
                    intValue = ((XlsBiffIntegerCell)cell).Value;
                    result.Value = TryConvertOADateTime(intValue, numberFormatIndex);
                    break;
                case BIFFRECORDTYPE.NUMBER:
                case BIFFRECORDTYPE.NUMBER_OLD:
                    doubleValue = ((XlsBiffNumberCell)cell).Value;
                    result.Value = TryConvertOADateTime(doubleValue, numberFormatIndex);
                    break;
                case BIFFRECORDTYPE.LABEL:
                case BIFFRECORDTYPE.LABEL_OLD:
                case BIFFRECORDTYPE.RSTRING:
                    result.Value = ((XlsBiffLabelCell)cell).GetValue(Encoding);
                    break;
                case BIFFRECORDTYPE.LABELSST:
                    result.Value = Workbook.SST.GetString(((XlsBiffLabelSSTCell)cell).SSTIndex, Encoding);
                    break;
                case BIFFRECORDTYPE.RK:
                    doubleValue = ((XlsBiffRKCell)cell).Value;
                    result.Value = TryConvertOADateTime(doubleValue, numberFormatIndex);
                    break;
                case BIFFRECORDTYPE.BLANK:
                case BIFFRECORDTYPE.BLANK_OLD:
                case BIFFRECORDTYPE.MULBLANK:
                    // Skip blank cells
                    break;
                case BIFFRECORDTYPE.FORMULA:
                case BIFFRECORDTYPE.FORMULA_V3:
                case BIFFRECORDTYPE.FORMULA_V4:
                    objectValue = TryGetFormulaValue(biffStream, (XlsBiffFormulaCell)cell, numberFormatIndex);
                    result.Value = objectValue;
                    break;
            }

            LogManager.Log(this).Debug("VALUE: {0}", result.Value);

            return result;
        }

        private object TryGetFormulaValue(XlsBiffStream biffStream, XlsBiffFormulaCell formulaCell, int numberFormatIndex)
        {
            switch (formulaCell.FormulaType)
            {
                case XlsBiffFormulaCell.FormulaValueType.Boolean:
                    return formulaCell.BooleanValue;
                case XlsBiffFormulaCell.FormulaValueType.Error:
                    return null;
                case XlsBiffFormulaCell.FormulaValueType.EmptyString:
                    return string.Empty;
                case XlsBiffFormulaCell.FormulaValueType.Number:
                    return TryConvertOADateTime(formulaCell.XNumValue, numberFormatIndex);
                case XlsBiffFormulaCell.FormulaValueType.String:
                    return TryGetFormulaString(biffStream);
            }

            // Bad data or new formula value type
            return null;
        }

        private string TryGetFormulaString(XlsBiffStream biffStream)
        {
            var rec = biffStream.Read();
            if (rec != null && rec.Id == BIFFRECORDTYPE.SHAREDFMLA)
            {
                rec = biffStream.Read();
            }

            if (rec != null && rec.Id == BIFFRECORDTYPE.STRING)
            {
                var stringRecord = (XlsBiffFormulaString)rec;
                return stringRecord.GetValue(Encoding);
            }

            // Bad data - could not find a string following the formula
            return null;
        }

        private object TryConvertOADateTime(double value, int numberFormatIndex)
        {
            var format = GetNumberFormatString(numberFormatIndex);
            if (format != null)
            {
                if (format.IsDateTimeFormat)
                    return Helpers.ConvertFromOATime(value, IsDate1904);
                if (format.IsTimeSpanFormat)
                    return TimeSpan.FromDays(value);
            }

            return value;
        }

        private object TryConvertOADateTime(int value, int numberFormatIndex)
        {
            var format = GetNumberFormatString(numberFormatIndex);
            if (format != null)
            {
                if (format.IsDateTimeFormat)
                    return Helpers.ConvertFromOATime(value, IsDate1904);
                if (format.IsTimeSpanFormat)
                    return TimeSpan.FromDays(value);
            }

            return value;
        }

        private int GetNumberFormatIndex(int xFormat)
        {
            int format;
            if (xFormat < ExtendedFormats.Count)
            {
                // If a cell XF record does not contain explicit attributes in a group (if the attribute group flag is not set),
                // it repeats the attributes of its style XF record.
                format = ExtendedFormats[xFormat].Format;
            }
            else
            {
                format = xFormat;
            }

            return format;
        }

        private void ReadWorksheetGlobals()
        {
            using (var biffStream = new XlsBiffStream(Stream, (int)DataOffset, Workbook.BiffVersion, null, Workbook.SecretKey, Workbook.Encryption))
            {
                // Check the expected BOF record was found in the BIFF stream
                if (biffStream.BiffVersion == 0 || biffStream.BiffType != BIFFTYPE.Worksheet)
                    return;

                XlsBiffHeaderFooterString header = null;
                XlsBiffHeaderFooterString footer = null;

                int maxCellColumn = 0;
                int maxRowCount = 0;

                var mergeCells = new List<CellRange>();
                var biffFormats = new Dictionary<ushort, XlsBiffFormatString>();
                var recordOffset = biffStream.Position;
                var rec = biffStream.Read();

                while (rec != null && !(rec is XlsBiffEof))
                {
                    if (rec is XlsBiffDimensions dims)
                    {
                        FieldCount = dims.LastColumn;
                        RowCount = (int)dims.LastRow;
                    }

                    if (rec.Id == BIFFRECORDTYPE.DEFAULTROWHEIGHT || rec.Id == BIFFRECORDTYPE.DEFAULTROWHEIGHT_V2)
                    {
                        var defaultRowHeightRecord = (XlsBiffDefaultRowHeight)rec;
                        DefaultRowHeight = defaultRowHeightRecord.RowHeight;
                    }

                    if (rec.Id == BIFFRECORDTYPE.RECORD1904)
                    {
                        IsDate1904 = ((XlsBiffSimpleValueRecord)rec).Value == 1;
                    }

                    if (rec.Id == BIFFRECORDTYPE.XF_V2 || rec.Id == BIFFRECORDTYPE.XF_V3 || rec.Id == BIFFRECORDTYPE.XF_V4)
                    {
                        ExtendedFormats.Add((XlsBiffXF)rec);
                    }

                    if (rec.Id == BIFFRECORDTYPE.MERGECELLS)
                    {
                        mergeCells.AddRange(((XlsBiffMergeCells)rec).MergeCells);
                    }
                    
                    if (rec.Id == BIFFRECORDTYPE.FORMAT)
                    {
                        var fmt = (XlsBiffFormatString)rec;
                        if (Workbook.BiffVersion >= 5)
                        {
                            // fmt.Index exists on BIFF5+ only
                            biffFormats.Add(fmt.Index, fmt);
                        }
                        else
                        {
                            biffFormats.Add((ushort)biffFormats.Count, fmt);
                        }
                    }

                    if (rec.Id == BIFFRECORDTYPE.FORMAT_V23)
                    {
                        var fmt = (XlsBiffFormatString)rec;
                        biffFormats.Add((ushort)biffFormats.Count, fmt);
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

                    if (rec.Id == BIFFRECORDTYPE.ROW)
                    {
                        var rowRecord = (XlsBiffRow)rec;
                        SetMinMaxRowOffset(rowRecord.RowIndex, recordOffset);
                        maxRowCount = Math.Max(maxRowCount, rowRecord.RowIndex + 1);
                    }

                    if (rec.IsCell)
                    {
                        var cell = (XlsBiffBlankCell)rec;
                        SetMinMaxRowOffset(cell.RowIndex, recordOffset);
                        maxCellColumn = Math.Max(maxCellColumn, cell.ColumnIndex + 1);
                        maxRowCount = Math.Max(maxRowCount, cell.RowIndex + 1);
                    }

                    recordOffset = biffStream.Position;
                    rec = biffStream.Read();

                    // Stop if we find the start out a new substream. Not always that files have the required EOF before a substream BOF.
                    if (rec is XlsBiffBOF)
                        break;
                }

                if (header != null || footer != null)
                {
                    HeaderFooter = new HeaderFooter(false, false)
                    {
                        OddHeader = header?.GetValue(Encoding),
                        OddFooter = footer?.GetValue(Encoding),
                    };
                }

                foreach (var biffFormat in biffFormats)
                {
                    var formatString = biffFormat.Value.GetValue(Encoding);
                    Formats.Add(biffFormat.Key, new NumberFormatString(formatString));
                }

                if (mergeCells.Count > 0)
                    MergeCells = mergeCells.ToArray();

                if (FieldCount < maxCellColumn)
                    FieldCount = maxCellColumn;

                if (RowCount < maxRowCount)
                    RowCount = maxRowCount;
            }
        }

        private bool GetMinMaxOffsetsForRowBlock(int rowIndex, int rowCount, out int minOffset, out int maxOffset)
        {
            minOffset = int.MaxValue;
            maxOffset = int.MinValue;

            for (var i = 0; i < rowCount; i++)
            {
                if (RowMinMaxOffsets.TryGetValue(rowIndex + i, out var minMax))
                {
                    minOffset = Math.Min(minOffset, minMax.Key);
                    maxOffset = Math.Max(maxOffset, minMax.Value);
                }
            }

            return minOffset != int.MaxValue;
        }

        private void SetMinMaxRowOffset(int rowIndex, int recordOffset)
        {
            if (!RowMinMaxOffsets.TryGetValue(rowIndex, out var minMax))
                minMax = new KeyValuePair<int, int>(int.MaxValue, int.MinValue);

            RowMinMaxOffsets[rowIndex] = new KeyValuePair<int, int>(
                Math.Min(minMax.Key, recordOffset),
                Math.Max(minMax.Value, recordOffset));
        }

        internal class XlsRowBlock
        {
            public Dictionary<int, Row> Rows { get; set; }
        }
    }
}