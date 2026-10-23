namespace ExcelDataReader.Core.BinaryFormat;

internal sealed class XlsWorkbookPartReader(XlsBiffStream biffStream, IXlsWorkbookPartSink sink)
{
    public void Load()
    {
        Dictionary<int, XlsBiffFormatString> formats = [];
        XlsBiffSST? sharedStrings = null;

        while (biffStream.ReadMetadataRecord() is { } record && record is not XlsBiffEof)
        {
            switch (record)
            {
                case XlsBiffBoundSheet sheet:
                    if (sheet.Type == XlsBiffBoundSheet.SheetType.Worksheet)
                        sink.AddSheet(sheet);
                    break;
                case XlsBiffSimpleValueRecord codePage when record.Id == BIFFRECORDTYPE.CODEPAGE:
                    sink.SetCodePage(codePage.Value);
                    break;
                case XlsBiffSimpleValueRecord is1904 when record.Id == BIFFRECORDTYPE.DATE1904:
                    sink.SetDate1904(is1904.Value == 1);
                    break;
                case XlsBiffFont font:
                    sink.AddFont(font);
                    break;
                case XlsBiffFormatString format23 when record.Id == BIFFRECORDTYPE.FORMAT_V23:
                    formats.Add((ushort)formats.Count, format23);
                    break;
                case XlsBiffFormatString format when record.Id == BIFFRECORDTYPE.FORMAT:
                    int index = format.Index;
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
                    formats.TryAdd(index, format);
#else
                    if (!formats.ContainsKey(index))
                        formats.Add(index, format);
#endif
                    break;
                case XlsBiffXF xf:
                    sink.AddXf(xf);
                    break;
                case XlsBiffSST sst:
                    sharedStrings = sst;
                    sink.SetSharedStringTable(sst);
                    break;
                case XlsBiffContinue continuation:
                    sharedStrings?.ReadContinueStrings(continuation);
                    break;
                case { Id: BIFFRECORDTYPE.WINDOW1 }:
                    sink.SetActiveSheet(record.ReadInt16(10));
                    break;
            }
        }

        sharedStrings?.Flush();

        // Decode after the scan because the workbook code page may follow FORMAT records.
        foreach (var format in formats)
            sink.AddNumberFormat(format.Key, format.Value.GetValue(sink.Encoding));
    }
}
