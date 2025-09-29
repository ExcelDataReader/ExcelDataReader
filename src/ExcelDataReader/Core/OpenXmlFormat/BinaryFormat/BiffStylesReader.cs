﻿using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;

internal sealed class BiffStylesReader(Stream stream) : BiffReader(stream)
{
    private const int Xf = 0x2f;

    private const int CellXfStart = 0x269;
    private const int CellXfEnd = 0x26a;

    private const int CellStyleXfStart = 0x272;
    private const int CellStyleXfEnd = 0x273;

    private const int NumberFormatStart = 0x267;
    private const int NumberFormat = 0x2c;
    private const int NumberFormatEnd = 0x268;

    private bool _inCellXf;
    private bool _inCellStyleXf;
    private bool _inNumberFormat;

    protected override Record ReadOverride(byte[] buffer, uint recordId, uint recordLength)
    {
        switch (recordId)
        {
            case CellXfStart:
                _inCellXf = true;
                break;
            case CellXfEnd:
                _inCellXf = false;
                break;
            case CellStyleXfStart:
                _inCellStyleXf = true;
                break;
            case CellStyleXfEnd:
                _inCellStyleXf = false;
                break;
            case NumberFormatStart:
                _inNumberFormat = true;
                break;
            case NumberFormatEnd:
                _inNumberFormat = false;
                break;

            case Xf when _inCellXf:
            case Xf when _inCellStyleXf:
                {
                    // var flags = buffer[14];
                    var extendedFormat = new ExtendedFormat(
                        GetWord(buffer, 0),
                        GetWord(buffer, 4),
                        GetWord(buffer, 2),
                        (buffer[13] & 0x10000) != 0,
                        (buffer[13] & 0x100000) != 0,
                        (int)(uint)buffer[11],
                        (HorizontalAlignment)(buffer[12] & 0b111),
                        (VerticalAlignment)(buffer[12] >> 3 & 0b111));

                    if (_inCellXf)
                        return new ExtendedFormatRecord(extendedFormat);
                    return new CellStyleExtendedFormatRecord(extendedFormat);
                }

            case NumberFormat when _inNumberFormat:
                {
                    // Must be between 1 and 255 characters
                    int format = GetWord(buffer, 0);
                    uint length = GetDWord(buffer, 2);
                    string formatString = GetString(buffer, 2 + 4, length);

                    return new NumberFormatRecord(format, formatString);
                }
        }

        return Record.Default;
    }
}
