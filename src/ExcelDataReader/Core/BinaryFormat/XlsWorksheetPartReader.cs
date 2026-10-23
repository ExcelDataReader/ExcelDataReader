using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

internal sealed class XlsWorksheetPartReader(
    XlsWorkbook workbook,
    XlsBiffBoundSheet sheet,
    Stream stream,
    bool singlePassMode)
{
    public XlsWorksheetMetadata Load()
    {
        var metadata = new XlsWorksheetMetadata
        {
            Encoding = workbook.Encoding,
            IsDate1904 = workbook.IsDate1904,
        };

        using var biffStream = new XlsBiffStream(stream, (int)sheet.StartOffset, workbook.BiffVersion, BIFFTYPE.Worksheet, secretKey: workbook.SecretKey, encryption: workbook.Encryption);

        if (biffStream.BiffVersion == 0 || (biffStream.BiffType != BIFFTYPE.Worksheet && biffStream.BiffType != BIFFTYPE.MacroSheet))
            return metadata;

        XlsBiffHeaderFooterString? header = null;
        XlsBiffHeaderFooterString? footer = null;
        var ixfeOffset = -1;
        int maxCellColumn = 0;
        int maxRowCount = 0;
        int maxRowCountFromRowRecord = 0;
        List<CellRange> mergeCells = [];
        Dictionary<ushort, XlsBiffFormatString> biffFormats = [];
        List<Column> columnWidths = [];

        var stopAtCells = false;
        var isFirstRecord = true;
        while (!stopAtCells)
        {
            var recordOffset = biffStream.Position;
            if (!biffStream.ReadWorksheetRecord(out var item))
                break;

            try
            {
                if (item.IsCell)
                {
                    // Indexing only reads cell coordinates, never values or formula strings.
                    if (singlePassMode)
                        break;

                    isFirstRecord = false;
                    var frame = item.Cell;
                    maxCellColumn = Math.Max(
                        maxCellColumn,
                        XlsWorksheetCellDecoder.GetLastDefinedColumn(frame, item.OwnedRecord as XlsBiffBlankCell) + 1);

                    maxRowCount = Math.Max(maxRowCount, frame.RowIndex + 1);
                    if (ixfeOffset != -1)
                    {
                        SetMinMaxRowOffset(metadata, frame.RowIndex, ixfeOffset, maxRowCount - 1);
                        ixfeOffset = -1;
                    }

                    SetMinMaxRowOffset(metadata, frame.RowIndex, recordOffset, maxRowCount - 1);
                    continue;
                }

                var record = item.OwnedRecord!;
                if (XlsBiffStream.IsIgnoredMetadataRecord(record))
                    continue;

                if (record is XlsBiffEof)
                    break;

                // Some workbooks omit EOF before beginning another substream.
                if (record is XlsBiffBOF && !isFirstRecord)
                    break;

                isFirstRecord = false;
                switch (record)
                {
                    case XlsBiffDimensions dimensions:
                        metadata.RowCount = (int)dimensions.LastRow;
                        metadata.Dimension = new CellRange(dimensions.FirstColumn, (int)dimensions.FirstRow, dimensions.LastColumn - 1, (int)dimensions.LastRow - 1);
                        break;
                    case XlsBiffDefaultRowHeight defaultRowHeight:
                        metadata.DefaultRowHeight = defaultRowHeight.RowHeight;
                        break;
                    case XlsBiffSimpleValueRecord is1904 when record.Id == BIFFRECORDTYPE.DATE1904:
                        metadata.IsDate1904 = is1904.Value == 1;
                        break;
                    case XlsBiffXF xf when record.Id is BIFFRECORDTYPE.XF_V2 or BIFFRECORDTYPE.XF_V3 or BIFFRECORDTYPE.XF_V4:
                        // These XF records occur in raw BIFF2-4 sheets rather than workbook globals.
                        workbook.AddXf(xf);
                        break;
                    case XlsBiffMergeCells mergeCell:
                        mergeCells.AddRange(mergeCell.MergeCells);
                        break;
                    case XlsBiffColInfo columnInfo:
                        columnWidths.Add(columnInfo.Value);
                        break;
                    case XlsBiffFormatString format when record.Id == BIFFRECORDTYPE.FORMAT:
                        if (workbook.BiffVersion >= 5)
                        {
                            biffFormats.Add(format.Index, format);
                        }
                        else
                        {
                            biffFormats.Add((ushort)biffFormats.Count, format);
                        }

                        break;
                    case XlsBiffFormatString format23 when record.Id == BIFFRECORDTYPE.FORMAT_V23:
                        biffFormats.Add((ushort)biffFormats.Count, format23);
                        break;
                    case XlsBiffSimpleValueRecord codePage when record.Id == BIFFRECORDTYPE.CODEPAGE:
                        metadata.Encoding = EncodingHelper.GetEncoding(codePage.Value) ?? metadata.Encoding;
                        break;
                    case XlsBiffHeaderFooterString headerRecord when record.Id == BIFFRECORDTYPE.HEADER && record.RecordSize > 0:
                        header = headerRecord;
                        break;
                    case XlsBiffHeaderFooterString footerRecord when record.Id == BIFFRECORDTYPE.FOOTER && record.RecordSize > 0:
                        footer = footerRecord;
                        break;
                    case XlsBiffCodeName codeName:
                        metadata.CodeName = codeName.GetValue(metadata.Encoding);
                        break;
                    case XlsBiffRow row:
                        if (singlePassMode)
                        {
                            // Row and cell records are left for the forward-only reader.
                            stopAtCells = true;
                            break;
                        }

                        SetMinMaxRow(metadata, row.RowIndex, row);

                        // Row records affect RowCount but not the cell-overlap ranges.
                        maxRowCountFromRowRecord = Math.Max(maxRowCountFromRowRecord, row.RowIndex + 1);
                        break;
                    case { Id: BIFFRECORDTYPE.IXFE }:
                        ixfeOffset = recordOffset;
                        break;
                }
            }
            finally
            {
                item.Release();
            }
        }

        if (header != null || footer != null)
            metadata.HeaderFooter = new HeaderFooter(footer?.GetValue(metadata.Encoding), header?.GetValue(metadata.Encoding));

        foreach (var format in biffFormats)
            workbook.AddNumberFormat(format.Key, format.Value.GetValue(metadata.Encoding));

        metadata.MergeCells = mergeCells.ToArray();
        if (!singlePassMode)
            metadata.FieldCount = maxCellColumn;

        maxRowCount = Math.Max(maxRowCount, maxRowCountFromRowRecord);
        if (metadata.RowCount < maxRowCount)
            metadata.RowCount = maxRowCount;

        if (columnWidths.Count > 0)
            metadata.ColumnWidths = columnWidths;

        return metadata;
    }

    private static void SetMinMaxRow(XlsWorksheetMetadata metadata, int rowIndex, XlsBiffRow row)
    {
        if (!metadata.RowOffsetMap.TryGetValue(rowIndex, out var rowOffset))
        {
            rowOffset = new()
            {
                MinCellOffset = int.MaxValue,
                MaxCellOffset = int.MinValue,
                MaxOverlapRowIndex = int.MinValue,
            };
            metadata.RowOffsetMap.Add(rowIndex, rowOffset);
        }

        rowOffset.UseDefaultRowHeight = row.UseDefaultRowHeight;
        rowOffset.RowHeight = row.RowHeight;
    }

    private static void SetMinMaxRowOffset(XlsWorksheetMetadata metadata, int rowIndex, int recordOffset, int maxOverlapRow)
    {
        if (!metadata.RowOffsetMap.TryGetValue(rowIndex, out var rowOffset))
        {
            rowOffset = new()
            {
                MinCellOffset = int.MaxValue,
                MaxCellOffset = int.MinValue,
                MaxOverlapRowIndex = int.MinValue,
            };
            metadata.RowOffsetMap.Add(rowIndex, rowOffset);
        }

        rowOffset.MinCellOffset = Math.Min(recordOffset, rowOffset.MinCellOffset);
        rowOffset.MaxCellOffset = Math.Max(recordOffset, rowOffset.MaxCellOffset);
        rowOffset.MaxOverlapRowIndex = Math.Max(maxOverlapRow, rowOffset.MaxOverlapRowIndex);
    }
}