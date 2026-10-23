namespace ExcelDataReader.Core.BinaryFormat;

// Both scheduling policies consume frames here before the transport can reuse their bytes.
internal sealed class XlsWorksheetCellDecoder(XlsWorksheet worksheet)
{
    public static int GetLastDefinedColumn(XlsCellFrame frame, XlsBiffBlankCell? ownedCell) =>
        ownedCell is { IsEmpty: true } ? -1 : frame.Id == BIFFRECORDTYPE.MULRK ? frame.LastColumnIndex : frame.ColumnIndex;

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    public Cell Decode(XlsBiffStream stream, XlsCellFrame frame, XlsBiffBlankCell? ownedCell, ushort? ixfe, int column)
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
            _ => DecodeOtherCell(stream, ownedCell!, effectiveStyle, out error),
        };
        return worksheet.Workbook.CreateCell(column, value, effectiveStyle, error, worksheet.IsDate1904);
    }

    private DecodedCellValue DecodeOtherCell(XlsBiffStream stream, XlsBiffBlankCell cell, ExtendedFormat style, out CellError? error)
    {
        DecodedCellValue value = default;
        error = null;
        switch (cell.Id)
        {
            case BIFFRECORDTYPE.BOOLERR:
                if (cell.ReadByte(7) == 0)
                    value = new DecodedCellValue(cell.ReadByte(6) != 0);
                else
                    error = (CellError)cell.ReadByte(6);
                break;
            case BIFFRECORDTYPE.BOOLERR_OLD:
                if (cell.ReadByte(8) == 0)
                    value = new DecodedCellValue(cell.ReadByte(7) != 0);
                else
                    error = (CellError)cell.ReadByte(7);
                break;
            case BIFFRECORDTYPE.INTEGER:
            case BIFFRECORDTYPE.INTEGER_OLD:
                value = new DecodedCellValue(((XlsBiffIntegerCell)cell).Value);
                break;
            case BIFFRECORDTYPE.LABEL:
            case BIFFRECORDTYPE.LABEL_V2:
            case BIFFRECORDTYPE.RSTRING:
                value = new DecodedCellValue(((XlsBiffLabelCell)cell).GetValue(GetEncoding(style)));
                break;
            case BIFFRECORDTYPE.FORMULA:
            case BIFFRECORDTYPE.FORMULA_V3:
            case BIFFRECORDTYPE.FORMULA_V4:
                value = new DecodedCellValue(GetFormulaValue(stream, (XlsBiffFormulaCell)cell, style, out error));
                break;
        }

        return value;
    }

    private object? GetFormulaValue(XlsBiffStream stream, XlsBiffFormulaCell cell, ExtendedFormat style, out CellError? error)
    {
        error = null;
        switch (cell.FormulaType)
        {
            case XlsBiffFormulaCell.FormulaValueType.Boolean:
                return cell.BooleanValue;
            case XlsBiffFormulaCell.FormulaValueType.Error:
                error = cell.ErrorValue;
                return null;
            case XlsBiffFormulaCell.FormulaValueType.EmptyString:
                return string.Empty;
            case XlsBiffFormulaCell.FormulaValueType.Number:
                return cell.XNumValue;
            case XlsBiffFormulaCell.FormulaValueType.String:
                var text = stream.ReadFormulaStringRecord();
                if (text == null)
                    return null;

                try
                {
                    return text.GetValue(GetEncoding(style));
                }
                finally
                {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
                    text.Return();
#endif
                }

            default:
                return null;
        }
    }

    private System.Text.Encoding GetEncoding(ExtendedFormat style)
    {
        var index = style.FontIndex;
        return (index >= 0 && index < worksheet.Workbook.Fonts.Count
            ? worksheet.Workbook.Fonts[index].ByteStringEncoding
            : null) ?? worksheet.Encoding;
    }

    private int GetXfIndex(XlsCellFrame cell, ushort? ixfe)
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