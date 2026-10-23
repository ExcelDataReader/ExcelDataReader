namespace ExcelDataReader.Core.BinaryFormat;

internal sealed class XlsWorkbookPartReader(XlsBiffStream biffStream, int biffVersion, IXlsWorkbookPartSink sink, IStylesPartSink styles, ISharedStringStore stringStore)
{
    public void Load()
    {
        Dictionary<int, XlsStringData> formats = [];
        XlsSSTReader? sharedStrings = null;
        uint uniqueCount = 0;

        while (biffStream.ReadFrame(out var frame) && frame.Id != BIFFRECORDTYPE.EOF)
        {
            switch (frame.Id)
            {
                case BIFFRECORDTYPE.BOUNDSHEET:
                    var sheet = new XlsSheetInfo(frame, biffVersion);
                    if (sheet.Type == XlsSheetInfo.SheetType.Worksheet)
                        sink.AddSheet(sheet);
                    break;
                case BIFFRECORDTYPE.CODEPAGE:
                    sink.SetCodePage(frame.ReadUInt16(0));
                    break;
                case BIFFRECORDTYPE.DATE1904:
                    sink.SetDate1904(frame.ReadUInt16(0) == 1);
                    break;
                case BIFFRECORDTYPE.FONT:
                    sink.AddFont(new XlsFontInfo(frame, biffVersion));
                    break;
                case BIFFRECORDTYPE.FORMAT_V23:
                    formats.Add((ushort)formats.Count, ReadFormat(frame, biffVersion));
                    break;
                case BIFFRECORDTYPE.FORMAT:
                    int index = frame.ReadUInt16(0);
                    var format = ReadFormat(frame, biffVersion);
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
                    formats.TryAdd(index, format);
#else
                    if (!formats.ContainsKey(index))
                        formats.Add(index, format);
#endif
                    break;
                case BIFFRECORDTYPE.XF:
                case BIFFRECORDTYPE.XF_V2:
                case BIFFRECORDTYPE.XF_V3:
                case BIFFRECORDTYPE.XF_V4:
                    new XlsCellStyle(frame, biffVersion).AddTo(styles);
                    break;
                case BIFFRECORDTYPE.SST:
                    uniqueCount = frame.ReadUInt32(4);
                    sharedStrings = new XlsSSTReader();
                    sink.SetSharedStringTable();
                    sharedStrings.ReadStrings(frame, 12, false, stringStore, RemainingStrings());
                    break;
                case BIFFRECORDTYPE.CONTINUE:
                    if ((uint)stringStore.Count < uniqueCount)
                        sharedStrings?.ReadStrings(frame, 4, true, stringStore, RemainingStrings());
                    break;
                case BIFFRECORDTYPE.WINDOW1:
                    sink.SetActiveSheet(frame.ReadInt16(10));
                    break;
            }
        }

        sharedStrings?.Flush(stringStore);

        // Decode after the scan because the workbook code page may follow FORMAT records.
        foreach (var format in formats)
            styles.AddNumberFormat(format.Key, format.Value.GetValue(sink.Encoding));

        int RemainingStrings() => (int)Math.Min(int.MaxValue, uniqueCount > (uint)stringStore.Count ? uniqueCount - (uint)stringStore.Count : 0);
    }

    internal static XlsStringData ReadFormat(XlsBiffFrame frame, int version)
    {
        if (frame.Id == BIFFRECORDTYPE.FORMAT_V23)
            return XlsStringData.Read(frame, 4, true, false);
        if (version is >= 2 and <= 5)
            return XlsStringData.Read(frame, 6, true, false);
        if (version == 8)
            return XlsStringData.Read(frame, 6, false, true);
        throw new ArgumentException("Unexpected BIFF version " + version, nameof(version));
    }
}