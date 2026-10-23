using System.Text;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat;

// Both scheduling policies consume frames here before the transport can reuse their bytes.
internal sealed class XlsWorksheetCellDecoder(XlsWorksheet worksheet)
{
    public static int GetLastDefinedColumn(XlsBiffFrame frame) =>
        frame.IsEmptyCell ? -1 : frame.Id == BIFFRECORDTYPE.MULRK ? frame.LastColumnIndex : frame.ColumnIndex;

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    public Cell Decode(XlsBiffStream stream, XlsBiffFrame frame, ushort? ixfe, int column)
    {
        if (frame.Id == BIFFRECORDTYPE.MULRK)
        {
            var style = worksheet.Workbook.GetEffectiveCellStyle(frame.GetMultiCellStyle((ushort)column), frame.Format);
            return worksheet.Workbook.CreateCell(column, new DecodedCellValue(frame.GetMultiCellValue((ushort)column)), style, null, worksheet.IsDate1904);
        }

        var effectiveStyle = worksheet.Workbook.GetEffectiveCellStyle(GetXfIndex(frame, ixfe), frame.Format);
        CellError? error = null;
        var value = frame.Id switch
        {
            BIFFRECORDTYPE.NUMBER or BIFFRECORDTYPE.NUMBER_OLD => new DecodedCellValue(frame.NumberValue),
            BIFFRECORDTYPE.LABELSST => DecodedCellValue.SharedString(frame.SharedStringIndex),
            BIFFRECORDTYPE.RK => new DecodedCellValue(frame.RkValue),
            BIFFRECORDTYPE.BOOLERR => DecodeBoolErr(frame, 6, 7, out error),
            BIFFRECORDTYPE.BOOLERR_OLD => DecodeBoolErr(frame, 7, 8, out error),
            BIFFRECORDTYPE.INTEGER or BIFFRECORDTYPE.INTEGER_OLD => new DecodedCellValue((int)frame.ReadUInt16(frame.Id == BIFFRECORDTYPE.INTEGER_OLD ? 7 : 6)),
            BIFFRECORDTYPE.LABEL or BIFFRECORDTYPE.LABEL_V2 or BIFFRECORDTYPE.RSTRING =>
                new DecodedCellValue(GetLabelValue(frame, GetEncoding(effectiveStyle))),
            BIFFRECORDTYPE.FORMULA or BIFFRECORDTYPE.FORMULA_V3 or BIFFRECORDTYPE.FORMULA_V4 =>
                new DecodedCellValue(GetFormulaValue(stream, frame, out error, effectiveStyle)),
            _ => default,
        };
        return worksheet.Workbook.CreateCell(column, value, effectiveStyle, error, worksheet.IsDate1904);
    }

    private static DecodedCellValue DecodeBoolErr(XlsBiffFrame frame, int valueOffset, int typeOffset, out CellError? error)
    {
        error = null;
        var value = frame.ReadByte(valueOffset);
        if (frame.ReadByte(typeOffset) == 0)
            return new DecodedCellValue(value != 0);

        error = (CellError)value;
        return default;
    }

    private object? GetFormulaValue(XlsBiffStream stream, XlsBiffFrame frame, out CellError? error, ExtendedFormat style)
    {
        error = null;
        if (worksheet.Workbook.BiffVersion == 2)
            return frame.ReadDouble(7);

        if (frame.ReadUInt16(12) != 0xFFFF)
            return frame.ReadDouble(6);

        var formulaValueByte1 = frame.ReadByte(6);
        var formulaValueByte3 = frame.ReadByte(8);
        switch (formulaValueByte1)
        {
            case 0x00:
                return XlsFormulaStringReader.Read(stream, worksheet.Workbook.BiffVersion, GetEncoding(style));

            case 0x01:
                return formulaValueByte3 != 0;
            case 0x02:
                error = (CellError)formulaValueByte3;
                return null;
            case 0x03:
                return string.Empty;
            default:
                return null;
        }
    }

    private string GetLabelValue(XlsBiffFrame frame, Encoding encoding)
    {
        int biffVersion = worksheet.Workbook.BiffVersion;
        int offset;
        int characterCount;
        bool unicode;
        bool multiByte;
        switch (frame.Id)
        {
            case BIFFRECORDTYPE.LABEL_V2:
                offset = 11;
                characterCount = frame.ReadStringByteAt(offset);
                unicode = false;
                multiByte = false;
                break;
            case BIFFRECORDTYPE.LABEL when biffVersion is >= 2 and <= 5:
            case BIFFRECORDTYPE.RSTRING when biffVersion is >= 2 and <= 5:
                offset = 10;
                characterCount = frame.ReadUInt16At(offset);
                unicode = false;
                multiByte = false;
                break;
            case BIFFRECORDTYPE.LABEL when biffVersion == 8:
            case BIFFRECORDTYPE.RSTRING when biffVersion == 8:
                offset = 10;
                characterCount = frame.ReadUInt16At(offset);
                unicode = true;
                multiByte = false;
                break;
            default:
                throw new ArgumentException("Unexpected BIFF version " + biffVersion);
        }

        if (unicode)
        {
            if (characterCount == 0)
                return string.Empty;

            multiByte = (frame.ReadStringByteAt(offset + 2) & 1) != 0;
            int start = offset + 3;
            int byteCount = multiByte ? characterCount * 2 : characterCount;
            if (frame.RecordSize + 4 - start < byteCount)
                throw new ExcelReaderException(Errors.ErrorBiffStringSize);

            if (multiByte)
                return Encoding.Unicode.GetString(frame.Bytes, start, byteCount);

#if NET8_0_OR_GREATER
            return Encoding.Latin1.GetString(frame.Bytes.AsSpan(start, characterCount));
#elif NETSTANDARD2_1_OR_GREATER
            return string.Create(characterCount, (Bytes: frame.Bytes, Start: start), static (span, state) =>
            {
                for (int i = 0; i < span.Length; i++)
                    span[i] = (char)state.Bytes[state.Start + i];
            });
#else
            var characters = new char[characterCount];
            for (int i = 0; i < characterCount; i++)
                characters[i] = (char)frame.Bytes[start + i];
            return new string(characters);
#endif
        }

        offset += frame.Id == BIFFRECORDTYPE.LABEL_V2 ? 1 : 2;
        int count = characterCount * (Helpers.IsSingleByteEncoding(encoding) ? 1 : 2);
        frame.EnsureRange(offset, count);
        return encoding.GetString(frame.Bytes, offset, count);
    }

    private System.Text.Encoding GetEncoding(ExtendedFormat style)
    {
        var index = style.FontIndex;
        return (index >= 0 && index < worksheet.Workbook.Fonts.Count
            ? worksheet.Workbook.Fonts[index].ByteStringEncoding
            : null) ?? worksheet.Encoding;
    }

    private int GetXfIndex(XlsBiffFrame cell, ushort? ixfe)
    {
        if (worksheet.Workbook.BiffVersion != 2)
            return cell.XFormat;

        return cell.XFormat switch
        {
            63 when ixfe != null => ixfe.Value,
            > 63 => -1,
            { } when cell.XFormat < worksheet.Workbook.ExtendedFormats.Count => cell.XFormat,
            _ => -1,
        };
    }
}