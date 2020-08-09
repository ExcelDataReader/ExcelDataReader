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

            VisibleState = refSheet.VisibleState switch
            {
                XlsBiffBoundSheet.SheetVisibility.Hidden => "hidden",
                XlsBiffBoundSheet.SheetVisibility.VeryHidden => "veryhidden",
                _ => "visible",
            };
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

        public Column[] ColumnWidths { get; private set; }

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
                        yield return new Row(rowIndex, DefaultRowHeight / 20.0, new List<Cell>());
                    }

                    rowIndex++;
                    yield return rowBlock;
                }
            }
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
                if (RowOffsetMap.TryGetValue(startRow + i, out _))
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

                if (rec is XlsBiffBlankCell cell)
                {
                    var currentRow = EnsureRow(result, cell.RowIndex);

                    if (cell.Id == BIFFRECORDTYPE.MULRK)
                    {
                        var cellValues = ReadMultiCell(cell);
                        currentRow.Cells.AddRange(cellValues);
                    }
                    else
                    {
                        var xfIndex = GetXfIndexForCell(cell, ixfe);
                        var cellValue = ReadSingleCell(biffStream, cell, xfIndex);
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

                currentRow = new Row(rowIndex, height, new List<Cell>());

                result.Rows.Add(rowIndex, currentRow);
            }

            return currentRow;
        }

        private IEnumerable<Cell> ReadMultiCell(XlsBiffBlankCell cell)
        {
            LogManager.Log(this).Debug("ReadMultiCell {0}", cell.Id);

            switch (cell.Id)
            {
                case BIFFRECORDTYPE.MULRK:

                    XlsBiffMulRKCell rkCell = (XlsBiffMulRKCell)cell;
                    ushort lastColumnIndex = rkCell.LastColumnIndex;
                    for (ushort j = cell.ColumnIndex; j <= lastColumnIndex; j++)
                    {
                        var xfIndex = rkCell.GetXF(j);
                        var effectiveStyle = Workbook.GetEffectiveCellStyle(xfIndex, cell.Format);

                        var value = TryConvertOADateTime(rkCell.GetValue(j), effectiveStyle.NumberFormatIndex);
                        LogManager.Log(this).Debug("CELL[{0}] = {1}", j, value);
                        yield return new Cell(j, value, effectiveStyle, null);
                    }

                    break;
            }
        }

        /// <summary>
        /// Reads additional records if needed: a string record might follow a formula result
        /// </summary>
        private Cell ReadSingleCell(XlsBiffStream biffStream, XlsBiffBlankCell cell, int xfIndex)
        {
            LogManager.Log(this).Debug("ReadSingleCell {0}", cell.Id);

            var effectiveStyle = Workbook.GetEffectiveCellStyle(xfIndex, cell.Format);
            var numberFormatIndex = effectiveStyle.NumberFormatIndex;

            object value = null;
            CellError? error = null;
            switch (cell.Id)
            {
                case BIFFRECORDTYPE.BOOLERR:
                    if (cell.ReadByte(7) == 0)
                        value = cell.ReadByte(6) != 0;
                    else
                        error = (CellError)cell.ReadByte(6);
                    break;
                case BIFFRECORDTYPE.BOOLERR_OLD:
                    if (cell.ReadByte(8) == 0)
                        value = cell.ReadByte(7) != 0;
                    else
                        error = (CellError)cell.ReadByte(7);
                    break;
                case BIFFRECORDTYPE.INTEGER:
                case BIFFRECORDTYPE.INTEGER_OLD:
                    value = TryConvertOADateTime(((XlsBiffIntegerCell)cell).Value, numberFormatIndex);
                    break;
                case BIFFRECORDTYPE.NUMBER:
                case BIFFRECORDTYPE.NUMBER_OLD:
                    value = TryConvertOADateTime(((XlsBiffNumberCell)cell).Value, numberFormatIndex);
                    break;
                case BIFFRECORDTYPE.LABEL:
                case BIFFRECORDTYPE.LABEL_OLD:
                case BIFFRECORDTYPE.RSTRING:
                    value = GetLabelString((XlsBiffLabelCell)cell, effectiveStyle);
                    break;
                case BIFFRECORDTYPE.LABELSST:
                    value = Workbook.SST.GetString(((XlsBiffLabelSSTCell)cell).SSTIndex, Encoding);
                    break;
                case BIFFRECORDTYPE.RK:
                    value = TryConvertOADateTime(((XlsBiffRKCell)cell).Value, numberFormatIndex);
                    break;
                case BIFFRECORDTYPE.BLANK:
                case BIFFRECORDTYPE.BLANK_OLD:
                case BIFFRECORDTYPE.MULBLANK:
                    // Skip blank cells
                    break;
                case BIFFRECORDTYPE.FORMULA:
                case BIFFRECORDTYPE.FORMULA_V3:
                case BIFFRECORDTYPE.FORMULA_V4:
                    value = TryGetFormulaValue(biffStream, (XlsBiffFormulaCell)cell, effectiveStyle, out error);
                    break;
            }

            return new Cell(cell.ColumnIndex, value, effectiveStyle, error);
        }

        private string GetLabelString(XlsBiffLabelCell cell, ExtendedFormat effectiveStyle)
        {
            // 1. Use encoding from font's character set (BIFF5-8)
            // 2. If not specified, use encoding from CODEPAGE BIFF record
            // 3. If not specified, use configured fallback encoding
            // Encoding is only used on BIFF2-5 byte strings. BIFF8 uses XlsUnicodeString which ignores the encoding.
            var labelEncoding = GetFont(effectiveStyle.FontIndex)?.ByteStringEncoding ?? Encoding;
            return cell.GetValue(labelEncoding);
        }

        private XlsBiffFont GetFont(int fontIndex)
        {
            if (fontIndex < 0 || fontIndex >= Workbook.Fonts.Count)
            {
                return null;
            }

            return Workbook.Fonts[fontIndex];
        }

        private object TryGetFormulaValue(XlsBiffStream biffStream, XlsBiffFormulaCell formulaCell, ExtendedFormat effectiveStyle, out CellError? error)
        {
            error = null;
            switch (formulaCell.FormulaType)
            {
                case XlsBiffFormulaCell.FormulaValueType.Boolean: return formulaCell.BooleanValue;
                case XlsBiffFormulaCell.FormulaValueType.Error:
                    error = (CellError)formulaCell.ErrorValue;
                    return null;
                case XlsBiffFormulaCell.FormulaValueType.EmptyString: return string.Empty;
                case XlsBiffFormulaCell.FormulaValueType.Number: return TryConvertOADateTime(formulaCell.XNumValue, effectiveStyle.NumberFormatIndex);
                case XlsBiffFormulaCell.FormulaValueType.String: return TryGetFormulaString(biffStream, effectiveStyle);

                // Bad data or new formula value type
                default: return null;
            }
        }

        private string TryGetFormulaString(XlsBiffStream biffStream, ExtendedFormat effectiveStyle)
        {
            var rec = biffStream.Read();
            if (rec != null && rec.Id == BIFFRECORDTYPE.SHAREDFMLA)
            {
                rec = biffStream.Read();
            }

            if (rec != null && rec.Id == BIFFRECORDTYPE.STRING)
            {
                var stringRecord = (XlsBiffFormulaString)rec;
                var formulaEncoding = GetFont(effectiveStyle.FontIndex)?.ByteStringEncoding ?? Encoding; // Workbook.GetFontEncodingFromXF(xFormat) ?? Encoding;
                return stringRecord.GetValue(formulaEncoding);
            }

            // Bad data - could not find a string following the formula
            return null;
        }

        private object TryConvertOADateTime(double value, int numberFormatIndex)
        {
            var format = Workbook.GetNumberFormatString(numberFormatIndex);
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
            var format = Workbook.GetNumberFormatString(numberFormatIndex);
            if (format != null)
            {
                if (format.IsDateTimeFormat)
                    return Helpers.ConvertFromOATime(value, IsDate1904);
                if (format.IsTimeSpanFormat)
                    return TimeSpan.FromDays(value);
            }

            return value;
        }

        /// <summary>
        /// Returns an index into Workbook.ExtendedFormats for the given cell and preceding ixfe record.
        /// </summary>
        private int GetXfIndexForCell(XlsBiffBlankCell cell, XlsBiffRecord ixfe)
        {
            if (Workbook.BiffVersion == 2)
            {
                if (cell.XFormat == 63 && ixfe != null)
                {
                    var xFormat = ixfe.ReadUInt16(0);
                    return xFormat;
                }
                else if (cell.XFormat > 63)
                {
                    // Invalid XF ref on cell in BIFF2 stream
                    return -1;
                }
                else if (cell.XFormat < Workbook.ExtendedFormats.Count)
                {
                    return cell.XFormat;
                }
                else
                {
                    // Either the file has no XFs, or XF was out of range
                    return -1;
                }
            }

            return cell.XFormat;
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
                var columnWidths = new List<Column>();

                while (rec != null && !(rec is XlsBiffEof))
                {
                    switch (rec)
                    {
                        case XlsBiffDimensions dims:
                            FieldCount = dims.LastColumn;
                            RowCount = (int)dims.LastRow;
                            break;
                        case XlsBiffDefaultRowHeight defaultRowHeightRecord:
                            DefaultRowHeight = defaultRowHeightRecord.RowHeight;
                            break;
                        case XlsBiffSimpleValueRecord is1904 when rec.Id == BIFFRECORDTYPE.RECORD1904:
                            IsDate1904 = is1904.Value == 1;
                            break;
                        case XlsBiffXF xf when rec.Id == BIFFRECORDTYPE.XF_V2 || rec.Id == BIFFRECORDTYPE.XF_V3 || rec.Id == BIFFRECORDTYPE.XF_V4:
                            // NOTE: XF records should only occur in raw BIFF2-4 single worksheet documents without the workbook stream, or globally in the workbook stream.
                            // It is undefined behavior if multiple worksheets in a workbook declare XF records.
                            Workbook.AddXf(xf);
                            break;
                        case XlsBiffMergeCells mc:
                            mergeCells.AddRange(mc.MergeCells);
                            break;
                        case XlsBiffColInfo colInfo:
                            columnWidths.Add(colInfo.Value);
                            break;
                        case XlsBiffFormatString fmt when rec.Id == BIFFRECORDTYPE.FORMAT:
                            if (Workbook.BiffVersion >= 5)
                            {
                                // fmt.Index exists on BIFF5+ only
                                biffFormats.Add(fmt.Index, fmt);
                            }
                            else
                            {
                                biffFormats.Add((ushort)biffFormats.Count, fmt);
                            }

                            break;

                        case XlsBiffFormatString fmt23 when rec.Id == BIFFRECORDTYPE.FORMAT_V23:
                            biffFormats.Add((ushort)biffFormats.Count, fmt23);
                            break;
                        case XlsBiffSimpleValueRecord codePage when rec.Id == BIFFRECORDTYPE.CODEPAGE:
                            Encoding = EncodingHelper.GetEncoding(codePage.Value);
                            break;
                        case XlsBiffHeaderFooterString h when rec.Id == BIFFRECORDTYPE.HEADER && rec.RecordSize > 0:
                            header = h;
                            break;
                        case XlsBiffHeaderFooterString f when rec.Id == BIFFRECORDTYPE.FOOTER && rec.RecordSize > 0:
                            footer = f;
                            break;
                        case XlsBiffCodeName codeName:
                            CodeName = codeName.GetValue(Encoding);
                            break;
                        case XlsBiffRow row:
                            SetMinMaxRow(row.RowIndex, row);

                            // Count rows by row records without affecting the overlap in OffsetMap
                            maxRowCountFromRowRecord = Math.Max(maxRowCountFromRowRecord, row.RowIndex + 1);
                            break;
                        case XlsBiffBlankCell cell:
                            maxCellColumn = Math.Max(maxCellColumn, cell.ColumnIndex + 1);
                            maxRowCount = Math.Max(maxRowCount, cell.RowIndex + 1);
                            if (ixfeOffset != -1)
                            {
                                SetMinMaxRowOffset(cell.RowIndex, ixfeOffset, maxRowCount - 1);
                                ixfeOffset = -1;
                            }

                            SetMinMaxRowOffset(cell.RowIndex, recordOffset, maxRowCount - 1);
                            break;
                        case XlsBiffRecord ixfe when rec.Id == BIFFRECORDTYPE.IXFE:
                            ixfeOffset = recordOffset; 
                            break;
                    }

                    recordOffset = biffStream.Position;
                    rec = biffStream.Read();

                    // Stop if we find the start out a new substream. Not always that files have the required EOF before a substream BOF.
                    if (rec is XlsBiffBOF)
                        break;
                }

                if (header != null || footer != null)
                {
                    HeaderFooter = new HeaderFooter(footer?.GetValue(Encoding), header?.GetValue(Encoding));
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
