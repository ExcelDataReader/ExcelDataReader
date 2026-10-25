using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

internal sealed class XlsWorksheetPartReader(
    XlsWorkbook workbook,
    XlsSheetInfo sheet,
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

        using var biffStream = XlsBiffStartup.Open(stream, out var version, out var type, (int)sheet.StartOffset, workbook.BiffVersion, BIFFTYPE.Worksheet, secretKey: workbook.SecretKey, encryption: workbook.Encryption);

        if (version == 0 || (type != BIFFTYPE.Worksheet && type != BIFFTYPE.MacroSheet))
            return metadata;

        XlsStringData? header = null;
        XlsStringData? footer = null;
        var ixfeOffset = -1;
        int maxCellColumn = 0;
        int maxRowCount = 0;
        int maxRowCountFromRowRecord = 0;
        List<CellRange> mergeCells = [];
        Dictionary<ushort, XlsStringData> biffFormats = [];
        List<Column> columnWidths = [];

        var stopAtCells = false;
        var isFirstRecord = true;
        while (!stopAtCells)
        {
            var recordOffset = biffStream.Position;
            if (!biffStream.ReadFrame(out var frame))
                break;

            if (frame.IsCell)
            {
                // Indexing only reads cell coordinates, never values or formula strings.
                if (singlePassMode)
                    break;

                isFirstRecord = false;
                maxCellColumn = Math.Max(maxCellColumn, XlsWorksheetCellDecoder.GetLastDefinedColumn(frame) + 1);
                maxRowCount = Math.Max(maxRowCount, frame.RowIndex + 1);
                if (ixfeOffset != -1)
                {
                    SetMinMaxRowOffset(metadata, frame.RowIndex, ixfeOffset, maxRowCount - 1);
                    ixfeOffset = -1;
                }

                SetMinMaxRowOffset(metadata, frame.RowIndex, recordOffset, maxRowCount - 1);
                continue;
            }

            if (frame.Id == BIFFRECORDTYPE.EOF)
                break;

            // Some workbooks omit EOF before beginning another substream.
            if (frame.Id == BIFFRECORDTYPE.BOF && !isFirstRecord)
                break;

            isFirstRecord = false;
            if (frame.Id == BIFFRECORDTYPE.IXFE)
            {
                ixfeOffset = recordOffset;
                continue;
            }

            if (frame.Id is BIFFRECORDTYPE.ROW or BIFFRECORDTYPE.ROW_V2)
            {
                if (singlePassMode)
                {
                    // Row and cell records are left for the forward-only reader.
                    stopAtCells = true;
                    continue;
                }

                var rowIndex = frame.RowIndex;
                frame.ReadRowHeight(out bool useDefaultRowHeight, out int rowHeight);
                SetMinMaxRow(metadata, rowIndex, useDefaultRowHeight, rowHeight);
                maxRowCountFromRowRecord = Math.Max(maxRowCountFromRowRecord, rowIndex + 1);
                continue;
            }

            switch (frame.Id)
            {
                case BIFFRECORDTYPE.DIMENSIONS:
                case BIFFRECORDTYPE.DIMENSIONS_V2 when frame.RecordSize >= 8:
                    uint firstRow = workbook.BiffVersion < 8 ? frame.ReadUInt16(0) : frame.ReadUInt32(0);
                    uint lastRow = workbook.BiffVersion < 8 ? frame.ReadUInt16(2) : frame.ReadUInt32(4);
                    int columnOffset = workbook.BiffVersion < 8 ? 4 : 8;
                    metadata.RowCount = (int)lastRow;
                    metadata.Dimension = new CellRange(frame.ReadUInt16(columnOffset), (int)firstRow, frame.ReadUInt16(columnOffset + 2) - 1, (int)lastRow - 1);
                    break;
                case BIFFRECORDTYPE.DEFAULTROWHEIGHT:
                case BIFFRECORDTYPE.DEFAULTROWHEIGHT_V2:
                    metadata.DefaultRowHeight = workbook.BiffVersion == 2
                        ? frame.ReadUInt16(0) & 0x7FFF
                        : (frame.ReadUInt16(0) & 2) == 0 ? frame.ReadUInt16(2) : 0;
                    break;
                case BIFFRECORDTYPE.DATE1904:
                    metadata.IsDate1904 = frame.ReadUInt16(0) == 1;
                    break;
                case BIFFRECORDTYPE.XF_V2:
                case BIFFRECORDTYPE.XF_V3:
                case BIFFRECORDTYPE.XF_V4:
                    // These XF records occur in raw BIFF2-4 sheets rather than workbook globals.
                    new XlsCellStyle(frame, workbook.BiffVersion).AddTo(workbook);
                    break;
                case BIFFRECORDTYPE.MERGECELLS:
                    int count = frame.ReadUInt16(0);
                    for (int i = 0; i < count; i++)
                        mergeCells.Add(new CellRange(frame.ReadInt16(6 + i * 8), frame.ReadInt16(2 + i * 8), frame.ReadInt16(8 + i * 8), frame.ReadInt16(4 + i * 8)));
                    break;
                case BIFFRECORDTYPE.COLWIDTH:
                    columnWidths.Add(new Column(frame.ReadByte(0), frame.ReadByte(1), false, frame.ReadUInt16(2) / 256.0));
                    break;
                case BIFFRECORDTYPE.COLINFO:
                    var flags = frame.ReadUInt16(8);
                    columnWidths.Add(new Column(frame.ReadUInt16(0), frame.ReadUInt16(2), (flags & 1) != 0, (flags & 2) != 0 ? frame.ReadUInt16(4) / 256.0 : null));
                    break;
                case BIFFRECORDTYPE.FORMAT:
                    biffFormats.Add(workbook.BiffVersion >= 5 ? frame.ReadUInt16(0) : (ushort)biffFormats.Count, XlsWorkbookPartReader.ReadFormat(frame, workbook.BiffVersion));
                    break;
                case BIFFRECORDTYPE.FORMAT_V23:
                    biffFormats.Add((ushort)biffFormats.Count, XlsWorkbookPartReader.ReadFormat(frame, workbook.BiffVersion));
                    break;
                case BIFFRECORDTYPE.CODEPAGE:
                    metadata.Encoding = EncodingHelper.GetEncoding(frame.ReadUInt16(0)) ?? metadata.Encoding;
                    break;
                case BIFFRECORDTYPE.HEADER when frame.RecordSize > 0:
                    header = XlsStringData.Read(frame, 4, workbook.BiffVersion < 8, workbook.BiffVersion == 8);
                    break;
                case BIFFRECORDTYPE.FOOTER when frame.RecordSize > 0:
                    footer = XlsStringData.Read(frame, 4, workbook.BiffVersion < 8, workbook.BiffVersion == 8);
                    break;
                case BIFFRECORDTYPE.CODENAME:
                    metadata.CodeName = XlsStringData.Read(frame, 4, false, true).GetValue(metadata.Encoding);
                    break;
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

    private static void SetMinMaxRow(XlsWorksheetMetadata metadata, int rowIndex, bool useDefaultRowHeight, int rowHeight)
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

        rowOffset.UseDefaultRowHeight = useDefaultRowHeight;
        rowOffset.RowHeight = rowHeight;
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