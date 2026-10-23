namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;

internal sealed class BiffWorksheetReader(Stream stream, WorksheetReadMode mode, CommonWorkbook workbook, bool date1904 = false) : BiffReader(stream), IWorksheetReader
{
    private const uint Row = 0x00;
    private const uint Blank = 0x01;
    private const uint Number = 0x02;
    private const uint BoolError = 0x03;
    private const uint Bool = 0x04;
    private const uint Float = 0x05;
    private const uint String = 0x06;
    private const uint SharedString = 0x07;
    private const uint FormulaString = 0x08;
    private const uint FormulaNumber = 0x09;
    private const uint FormulaBool = 0x0a;
    private const uint FormulaError = 0x0b;

    private const uint BrtCellRString = 0x3E;
    private const uint SheetDataBegin = 0x91;
    private const uint SheetDataEnd = 0x92;
    private const uint SheetPr = 0x93;
    private const uint SheetDim = 0x94;
    private const uint SheetFormatPr = 0x1E5;
    private const uint Column = 0x3C;
    private const uint HeaderFooter = 0x1DF;
    private const uint MergeCell = 176;

    private int _scanMaxRowIndex = int.MinValue;
    private int _scanMaxColumnIndex = int.MinValue;
    private bool _inSheetData;
    private bool _rowOpen;
    private byte[]? _pendingBuffer;
    private uint _pendingRecordId;
    private bool _endOfInput;

    public WorksheetReadStatus Consume<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        while (!_endOfInput)
        {
            byte[] buffer;
            uint recordId;
            if (_pendingBuffer != null)
            {
                buffer = _pendingBuffer;
                recordId = _pendingRecordId;
                _pendingBuffer = null;
            }
            else if (!TryReadFrame(out buffer, out recordId, out _))
            {
                _endOfInput = true;
                if (_rowOpen)
                {
                    _rowOpen = false;
                    if (handler.OnRowEnd() == WorksheetReadControl.Pause)
                        return WorksheetReadStatus.Paused;
                }

                break;
            }

            if (_rowOpen && recordId is Row or SheetDataEnd)
            {
                _pendingBuffer = buffer;
                _pendingRecordId = recordId;
                _rowOpen = false;
                if (handler.OnRowEnd() == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
                continue;
            }

            WorksheetReadControl control = WorksheetReadControl.Continue;
            if (recordId == SheetDataBegin)
            {
                _inSheetData = true;
                control = handler.OnSheetDataBegin();
            }
            else if (recordId == SheetDataEnd)
            {
                bool scanCompleted = _inSheetData && mode == WorksheetReadMode.Scan;
                _inSheetData = false;
                if (scanCompleted)
                    control = handler.OnScanResult(_scanMaxRowIndex, _scanMaxColumnIndex);
            }
            else if (_inSheetData)
            {
                if (mode == WorksheetReadMode.Scan)
                    ScanRecord(buffer, recordId);

                if (recordId == Row)
                {
                    _rowOpen = true;
                    byte flags = buffer[11];
                    double? height = (flags & 0b100000) != 0 ? GetWord(buffer, 8) / 20.0 : null;
                    control = handler.OnRow(GetInt32(buffer, 0), (flags & 0b10000) != 0, height);
                }
                else if (mode == WorksheetReadMode.Materialize && !ConsumeCell(buffer, recordId, ref handler, out control))
                {
                    ReadMetadata(buffer, recordId, ref handler, out control);
                }
            }
            else if (!ReadMetadata(buffer, recordId, ref handler, out control))
            {
                // Preserve validation of stray cell records outside sheetData.
                var ignoringHandler = default(IgnoringHandler);
                ConsumeCell(buffer, recordId, ref ignoringHandler, out _);
            }

            if (control == WorksheetReadControl.Pause)
                return WorksheetReadStatus.Paused;
        }

        return WorksheetReadStatus.EndOfInput;
    }

    public override void Dispose()
    {
        _pendingBuffer = null;
        base.Dispose();
    }

    private static bool ReadMetadata<THandler>(byte[] buffer, uint recordId, ref THandler handler, out WorksheetReadControl control)
        where THandler : struct, IWorksheetHandler
    {
        switch (recordId)
        {
            case SheetPr:
                {
                    // Must be between 0 and 31 characters.
                    uint length = GetDWord(buffer, 19);
                    string? codeName = length == 0 ? null : GetString(buffer, 23, length);
                    control = handler.OnCodeName(codeName);
                    return true;
                }

            case SheetDim:
                {
                    int fromRow = GetInt32(buffer, 0);
                    int toRow = GetInt32(buffer, 4);
                    int fromColumn = GetInt32(buffer, 8);
                    int toColumn = GetInt32(buffer, 12);
                    control = handler.OnDimension(new CellRange(fromColumn, fromRow, toColumn, toRow));
                    return true;
                }

            case SheetFormatPr:
                {
                    double? defaultHeight = (buffer[8] & 0b1000) != 0 ? GetWord(buffer, 6) : null;
                    control = handler.OnDefaultRowHeight(defaultHeight);
                    return true;
                }

            case Column:
                {
                    int minimum = GetInt32(buffer, 0);
                    int maximum = GetInt32(buffer, 4);
                    byte flags = buffer[16];
                    double? width = (flags & 0b10) != 0 ? GetDWord(buffer, 8) / 256.0 : null;
                    control = handler.OnColumn(new Column(minimum, maximum, (flags & 0b1) != 0, width));
                    return true;
                }

            case HeaderFooter:
                {
                    byte flags = buffer[0];
                    bool differentOddEven = (flags & 1) != 0;
                    bool differentFirst = (flags & 0b10) != 0;
                    uint offset = 2;
                    var header = GetNullableString(buffer, ref offset);
                    var footer = GetNullableString(buffer, ref offset);
                    var headerEven = GetNullableString(buffer, ref offset);
                    var footerEven = GetNullableString(buffer, ref offset);
                    var headerFirst = GetNullableString(buffer, ref offset);
                    var footerFirst = GetNullableString(buffer, ref offset);
                    var value = new HeaderFooter(differentFirst, differentOddEven)
                    {
                        FirstHeader = headerFirst,
                        FirstFooter = footerFirst,
                        OddHeader = header,
                        OddFooter = footer,
                        EvenHeader = headerEven,
                        EvenFooter = footerEven,
                    };
                    control = handler.OnHeaderFooter(value);
                    return true;
                }

            case MergeCell:
                {
                    int fromRow = GetInt32(buffer, 0);
                    int toRow = GetInt32(buffer, 4);
                    int fromColumn = GetInt32(buffer, 8);
                    int toColumn = GetInt32(buffer, 12);
                    control = handler.OnMergeCell(new CellRange(fromColumn, fromRow, toColumn, toRow));
                    return true;
                }

            default:
                control = WorksheetReadControl.Continue;
                return false;
        }
    }

    private bool ConsumeCell<THandler>(byte[] buffer, uint recordId, ref THandler handler, out WorksheetReadControl control)
        where THandler : struct, IWorksheetHandler
    {
        control = WorksheetReadControl.Continue;
        DecodedCellValue value = default;
        CellError? error = null;
        switch (recordId)
        {
            case Blank:
                break;
            case BoolError:
            case FormulaError:
                error = (CellError)buffer[8];
                break;
            case Number:
                value = new DecodedCellValue(GetRkNumber(buffer, 8));
                break;
            case Bool:
            case FormulaBool:
                value = new DecodedCellValue(buffer[8] == 1);
                break;
            case FormulaNumber:
            case Float:
                value = new DecodedCellValue(GetDouble(buffer, 8));
                break;
            case String:
            case FormulaString:
                value = new DecodedCellValue(GetString(buffer, 12, GetDWord(buffer, 8)), CellValueKind.FormattedString);
                break;
            case BrtCellRString:
                value = new DecodedCellValue(GetString(buffer, 13, GetDWord(buffer, 9)), CellValueKind.FormattedString);
                break;
            case SharedString:
                value = DecodedCellValue.SharedString(GetDWord(buffer, 8));
                break;
            default:
                return false;
        }

        int style = (int)(GetDWord(buffer, 4) & 0xffffff);
        var cell = workbook.CreateCell((int)GetDWord(buffer, 0), value, workbook.GetEffectiveCellStyle(style, 0), error, date1904);
        control = handler.OnCell(in cell);
        return true;
    }

    private void ScanRecord(byte[] buffer, uint recordId)
    {
        if (recordId == Row)
        {
            int rowIndex = GetInt32(buffer, 0);
            if (rowIndex > _scanMaxRowIndex)
                _scanMaxRowIndex = rowIndex;

            return;
        }

        if (recordId is Blank)
            return;

        if (recordId is Number or Float or String or SharedString or FormulaNumber or FormulaString or FormulaBool or BrtCellRString or Bool or BoolError or FormulaError)
        {
            int column = (int)GetDWord(buffer, 0);
            if (column > _scanMaxColumnIndex)
                _scanMaxColumnIndex = column;
        }
    }

    private readonly struct IgnoringHandler : IWorksheetHandler
    {
        public WorksheetReadControl OnDimension(CellRange range) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnCodeName(string? codeName) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnDefaultRowHeight(double? height) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnColumn(Column column) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnMergeCell(CellRange range) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnHeaderFooter(HeaderFooter headerFooter) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnSheetDataBegin() => WorksheetReadControl.Continue;

        public WorksheetReadControl OnScanResult(int maxRowIndex, int maxColumnIndex) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnRow(int index, bool hidden, double? height) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnCell(in Cell cell) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnRowEnd() => WorksheetReadControl.Continue;
    }
}