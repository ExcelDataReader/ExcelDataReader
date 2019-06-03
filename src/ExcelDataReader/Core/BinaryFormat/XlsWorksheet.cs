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
            Encoding = workbook.Encoding;
            RowOffsetMap = new Dictionary<int, XlsRowOffset>();
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

        public Col[] ColumnWidths { get; private set; }

        /// <summary>
        /// Gets the worksheet data offset.
        /// </summary>
        public uint DataOffset { get; }

        public Stream Stream { get; }

        public Encoding Encoding { get; private set; }

        public double DefaultRowHeight { get; set; }

        public Dictionary<int, XlsRowOffset> RowOffsetMap { get; }

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
            if (Workbook.Formats.TryGetValue(numberFormatIndex, out var fmtString))
            {
                return fmtString;
            }

            return null;
        }

        /// <summary>
        /// Find how many rows to read at a time and their offset in the file.
        /// If rows are stored sequentially in the file, returns a block size of up to 32 rows.
        /// If rows are stored non-sequentially, the block size may extend up to the entire worksheet stream
        /// </summary>
        private void GetBlockSize(int startRow, out int blockRowCount, out int minOffset, out int maxOffset)
        {
            minOffset = int.MaxValue;
            maxOffset = int.MinValue;

            var i = 0;
            blockRowCount = Math.Min(32, RowCount - startRow);

            while (i < blockRowCount)
            {
                if (RowOffsetMap.TryGetValue(startRow + i, out var rowOffset))
                {
                    minOffset = Math.Min(rowOffset.MinCellOffset, minOffset);
                    maxOffset = Math.Max(rowOffset.MaxCellOffset, maxOffset);

                    if (rowOffset.MaxOverlapRowIndex != int.MinValue)
                    {
                        var maxOverlapRowCount = rowOffset.MaxOverlapRowIndex + 1;
                        blockRowCount = Math.Max(blockRowCount, maxOverlapRowCount - startRow);
                    }
                }

                i++;
            }
        }

        private IEnumerable<Row> ReadWorksheetRows(XlsBiffStream biffStream)
        {
            var rowIndex = 0;

            while (rowIndex < RowCount)
            {
                GetBlockSize(rowIndex, out var blockRowCount, out var minOffset, out var maxOffset);

                var block = ReadNextBlock(biffStream, rowIndex, blockRowCount, minOffset, maxOffset);
                
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

        private XlsRowBlock ReadNextBlock(XlsBiffStream biffStream, int startRow, int rows, int minOffset, int maxOffset)
        {
            var result = new XlsRowBlock { Rows = new Dictionary<int, Row>() };

            // Ensure rows with physical records are initialized with height
            for (var i = 0; i < rows; i++)
            {
                if (RowOffsetMap.TryGetValue(startRow + i, out var rowOffset))
                {
                    EnsureRow(result, startRow + i);
                }
            }

            if (minOffset == int.MaxValue)
            {
                return result;
            }

            biffStream.Position = minOffset;

            XlsBiffRecord rec;
            XlsBiffRecord ixfe = null;
            while (biffStream.Position <= maxOffset && (rec = biffStream.Read()) != null)
            {
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
                        var cellStyle = GetCellStyle(cell, ixfe);
                        var cellValue = ReadSingleCell(biffStream, cell, cellStyle);
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
                var height = DefaultRowHeight / 20.0;
                if (RowOffsetMap.TryGetValue(rowIndex, out var rowOffset) && rowOffset.Record != null)
                {
                    height = (rowOffset.Record.UseDefaultRowHeight ? DefaultRowHeight : rowOffset.Record.RowHeight) / 20.0;
                }

                currentRow = new Row()
                {
                    RowIndex = rowIndex,
                    Height = height,
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
                        CellStyle cellStyle = new CellStyle();
                        Workbook.GetCellStyleFromXF(cellStyle, rkCell.GetXF(j));
                        var numberFormatIndex = Workbook.GetNumberFormatFromFileIndex(cellStyle.FormatIndex);
                        var resultCell = new Cell
                        {
                            ColumnIndex = j,
                            Value = TryConvertOADateTime(rkCell.GetValue(j), numberFormatIndex),
                            NumberFormatIndex = numberFormatIndex, 
                            CellStyle = cellStyle,
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
        private Cell ReadSingleCell(XlsBiffStream biffStream, XlsBiffBlankCell cell, CellStyle cellStyle)
        {
            LogManager.Log(this).Debug("ReadSingleCell {0}", cell.Id);

            double doubleValue;
            int intValue;
            object objectValue;

            var result = new Cell()
            {
                ColumnIndex = cell.ColumnIndex,
                NumberFormatIndex = cellStyle.FormatIndex,
                CellStyle = cellStyle,
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
                    result.Value = TryConvertOADateTime(intValue, cellStyle.FormatIndex);
                    break;
                case BIFFRECORDTYPE.NUMBER:
                case BIFFRECORDTYPE.NUMBER_OLD:
                    doubleValue = ((XlsBiffNumberCell)cell).Value;
                    result.Value = TryConvertOADateTime(doubleValue, cellStyle.FormatIndex);
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
                    result.Value = TryConvertOADateTime(doubleValue, cellStyle.FormatIndex);
                    break;
                case BIFFRECORDTYPE.BLANK:
                case BIFFRECORDTYPE.BLANK_OLD:
                case BIFFRECORDTYPE.MULBLANK:
                    // Skip blank cells
                    break;
                case BIFFRECORDTYPE.FORMULA:
                case BIFFRECORDTYPE.FORMULA_V3:
                case BIFFRECORDTYPE.FORMULA_V4:
                    objectValue = TryGetFormulaValue(biffStream, (XlsBiffFormulaCell)cell, cellStyle.FormatIndex);
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
        
        private CellStyle GetCellStyle(XlsBiffBlankCell cell, XlsBiffRecord ixfe)
        {
            CellStyle cellStyle = new CellStyle();
            if (Workbook.BiffVersion == 2)
            {
                if (cell.XFormat == 63 && ixfe != null)
                {
                    var xFormat = ixfe.ReadUInt16(0);
                    cellStyle.FormatIndex = Workbook.GetNumberFormatFromXF(xFormat);
                }
                else if (cell.XFormat > 63)
                {
                    // Invalid XF ref on cell in BIFF2 stream, default to built-in "General"
                    return cellStyle;
                }
                else if (cell.XFormat < Workbook.GetExtendedFormatCount)
                {
                    cellStyle.FormatIndex = Workbook.GetNumberFormatFromXF(cell.XFormat);
                }
                else
                {
                    // Either the file has no XFs, or XF was out of range. Use the cell attributes' format reference.
                    cellStyle.FormatIndex = Workbook.GetNumberFormatFromFileIndex(cell.Format);
                }
            }
            else
            {
                Workbook.GetCellStyleFromXF(cellStyle, cell.XFormat);
            }

            return cellStyle;
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
                var ixfeOffset = -1;

                int maxCellColumn = 0;
                int maxRowCount = 0; // number of rows with cell records
                int maxRowCountFromRowRecord = 0; // number of rows with row records

                var mergeCells = new List<CellRange>();
                var biffFormats = new Dictionary<ushort, XlsBiffFormatString>();
                var recordOffset = biffStream.Position;
                var rec = biffStream.Read();
                var columnWidths = new List<Col>();

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

                    if (rec.Id == BIFFRECORDTYPE.XF_V2)
                    {
                        // NOTE: XF records should only occur in raw BIFF2-4 single worksheet documents without the workbook stream, or globally in the workbook stream.
                        // It is undefined behavior if multiple worksheets in a workbook declare XF records.
                        var xf = (XlsBiffXF)rec;
                        Workbook.AddExtendedFormat(
                            0, // Not applicable for biff2
                            true,
                            xf.Format,
                            true,
                            0, // Not applicable for biff2
                            xf.HorizontalAlignment);
                    }
                    else if (rec.Id == BIFFRECORDTYPE.XF_V3 || rec.Id == BIFFRECORDTYPE.XF_V4)
                    {
                        // NOTE: XF records should only occur in raw BIFF2-4 single worksheet documents without the workbook stream, or globally in the workbook stream.
                        // It is undefined behavior if multiple worksheets in a workbook declare XF records.
                        var xf = (XlsBiffXF)rec;
                        Workbook.AddExtendedFormat(
                            xf.Parent,
                            (xf.UsedAttributes & XfUsedAttributes.NumberFormat) != 0,
                            xf.Format,
                            (xf.UsedAttributes & XfUsedAttributes.TextStyle) != 0,
                            xf.IndentLevel,
                            xf.HorizontalAlignment);
                    }

                    if (rec.Id == BIFFRECORDTYPE.MERGECELLS)
                    {
                        mergeCells.AddRange(((XlsBiffMergeCells)rec).MergeCells);
                    }

                    if (rec.Id == BIFFRECORDTYPE.COLINFO)
                    {
                        columnWidths.Add(((XlsBiffColInfo)rec).Value);
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

                    if (rec.Id == BIFFRECORDTYPE.ROW || rec.Id == BIFFRECORDTYPE.ROW_V2)
                    {
                        var rowRecord = (XlsBiffRow)rec;
                        SetMinMaxRow(rowRecord.RowIndex, rowRecord);
                        
                        // Count rows by row records without affecting the overlap in OffsetMap
                        maxRowCountFromRowRecord = Math.Max(maxRowCountFromRowRecord, rowRecord.RowIndex + 1);
                    }

                    if (rec.Id == BIFFRECORDTYPE.IXFE)
                    {
                        ixfeOffset = recordOffset;
                    }

                    if (rec.IsCell)
                    {
                        var cell = (XlsBiffBlankCell)rec;
                        maxCellColumn = Math.Max(maxCellColumn, cell.ColumnIndex + 1);
                        maxRowCount = Math.Max(maxRowCount, cell.RowIndex + 1);

                        if (ixfeOffset != -1)
                        {
                            SetMinMaxRowOffset(cell.RowIndex, ixfeOffset, maxRowCount - 1);
                            ixfeOffset = -1;
                        }

                        SetMinMaxRowOffset(cell.RowIndex, recordOffset, maxRowCount - 1);
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
                    Workbook.AddNumberFormat(biffFormat.Key, biffFormat.Value.GetValue(Encoding));
                }

                if (mergeCells.Count > 0)
                    MergeCells = mergeCells.ToArray();

                if (FieldCount < maxCellColumn)
                    FieldCount = maxCellColumn;

                maxRowCount = Math.Max(maxRowCount, maxRowCountFromRowRecord);
                if (RowCount < maxRowCount)
                    RowCount = maxRowCount;

                if (columnWidths.Count > 0)
                {
                    ColumnWidths = columnWidths.ToArray();
                }
            }
        }

        private void SetMinMaxRow(int rowIndex, XlsBiffRow row)
        {
            if (!RowOffsetMap.TryGetValue(rowIndex, out var rowOffset))
            {
                rowOffset = new XlsRowOffset();
                rowOffset.MinCellOffset = int.MaxValue;
                rowOffset.MaxCellOffset = int.MinValue;
                rowOffset.MaxOverlapRowIndex = int.MinValue;
                RowOffsetMap.Add(rowIndex, rowOffset);
            }

            rowOffset.Record = row;
        }

        private void SetMinMaxRowOffset(int rowIndex, int recordOffset, int maxOverlapRow)
        {
            if (!RowOffsetMap.TryGetValue(rowIndex, out var rowOffset))
            {
                rowOffset = new XlsRowOffset();
                rowOffset.MinCellOffset = int.MaxValue;
                rowOffset.MaxCellOffset = int.MinValue;
                rowOffset.MaxOverlapRowIndex = int.MinValue;
                RowOffsetMap.Add(rowIndex, rowOffset);
            }

            rowOffset.MinCellOffset = Math.Min(recordOffset, rowOffset.MinCellOffset);
            rowOffset.MaxCellOffset = Math.Max(recordOffset, rowOffset.MaxCellOffset);
            rowOffset.MaxOverlapRowIndex = Math.Max(maxOverlapRow, rowOffset.MaxOverlapRowIndex);
        }

        internal class XlsRowBlock
        {
            public Dictionary<int, Row> Rows { get; set; }
        }

        internal class XlsRowOffset
        {
            public XlsBiffRow Record { get; set; }

            public int MinCellOffset { get; set; }

            public int MaxCellOffset { get; set; }

            public int MaxOverlapRowIndex { get; set; }
        }
    }
}