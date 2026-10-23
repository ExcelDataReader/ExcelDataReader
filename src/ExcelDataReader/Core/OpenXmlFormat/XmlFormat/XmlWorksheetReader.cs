using System.Globalization;
using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

internal sealed class XmlWorksheetReader(XmlReader reader, WorksheetReadMode mode, CommonWorkbook workbook, bool date1904 = false) : IWorksheetReader
{
    private const string NWorksheet = "worksheet";
    private const string NSheetData = "sheetData";
    private const string NRow = "row";
    private const string ARef = "ref";
    private const string AR = "r";
    private const string NV = "v";
    private const string NIs = "is";
    private const string AT = "t";
    private const string AS = "s";
    private const string NC = "c";
    private const string NInlineStr = "inlineStr";
    private const string NStr = "str";
    private const string NMergeCells = "mergeCells";
    private const string NSheetProperties = "sheetPr";
    private const string NSheetFormatProperties = "sheetFormatPr";
    private const string ADefaultRowHeight = "defaultRowHeight";
    private const string NHeaderFooter = "headerFooter";
    private const string ADifferentFirst = "differentFirst";
    private const string ADifferentOddEven = "differentOddEven";
    private const string NFirstHeader = "firstHeader";
    private const string NFirstFooter = "firstFooter";
    private const string NOddHeader = "oddHeader";
    private const string NOddFooter = "oddFooter";
    private const string NEvenHeader = "evenHeader";
    private const string NEvenFooter = "evenFooter";
    private const string NCols = "cols";
    private const string NCol = "col";
    private const string AMin = "min";
    private const string AMax = "max";
    private const string AHidden = "hidden";
    private const string AWidth = "width";
    private const string ACustomWidth = "customWidth";
    private const string NMergeCell = "mergeCell";
    private const string NDimension = "dimension";
    private const string AHt = "ht";

    private readonly XmlReader _reader = reader;
    private readonly XmlProperNamespaces _properNamespaces = new(reader.IsStartElement() && reader.NamespaceURI == XmlNamespaces.StrictNsSpreadsheetMl);
    private bool _metadataStarted;
    private bool _endOfInput;
    private int _scanMaxRowIndex = int.MinValue;
    private int _scanMaxColumnIndex = int.MinValue;
    private bool _inSheetData;
    private bool _pendingRowContents;
    private bool _inRowContents;
    private bool _rowOpen;
    private bool _pendingRowAdvance;
    private bool _pendingScanResult;
    private int _rowIndex = -1;
    private int _nextColumnIndex;

    public WorksheetReadStatus Consume<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        if (_endOfInput)
            return WorksheetReadStatus.EndOfInput;

        while (true)
        {
            if (_pendingScanResult)
            {
                _pendingScanResult = false;
                if (handler.OnScanResult(_scanMaxRowIndex, _scanMaxColumnIndex) == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
            }

            if (_inSheetData)
            {
                while (_inSheetData)
                {
                    if (_pendingRowAdvance)
                    {
                        _pendingRowAdvance = false;
                        _reader.Read();
                    }

                    if (_pendingRowContents)
                    {
                        _pendingRowContents = false;
                        _inRowContents = !_reader.IsEmptyElement && XmlReaderHelper.ReadFirstContent(_reader);
                        _nextColumnIndex = 0;
                    }

                    if (_inRowContents)
                    {
                        while (!_reader.EOF)
                        {
                            if (_reader.NodeType == XmlNodeType.Element && _reader.LocalName == NC)
                            {
                                var control = ConsumeCell(_nextColumnIndex, out int columnIndex, ref handler);
                                _nextColumnIndex = columnIndex + 1;
                                if (control == WorksheetReadControl.Pause)
                                    return WorksheetReadStatus.Paused;
                            }
                            else if (_reader.NodeType == XmlNodeType.EndElement && _reader.LocalName == NRow)
                            {
                                break;
                            }
                            else if (!XmlReaderHelper.SkipContent(_reader))
                            {
                                break;
                            }
                        }

                        _inRowContents = false;
                    }

                    if (_rowOpen)
                    {
                        _rowOpen = false;
                        _pendingRowAdvance = true;
                        if (handler.OnRowEnd() == WorksheetReadControl.Pause)
                            return WorksheetReadStatus.Paused;
                        continue;
                    }

                    bool foundRow = false;
                    while (!_reader.EOF)
                    {
                        if (_reader.NodeType == XmlNodeType.Element && _reader.LocalName == NRow)
                        {
                            if (int.TryParse(_reader.GetAttribute(AR), out int arValue))
                                _rowIndex = arValue - 1;
                            else
                                _rowIndex++;

                            _scanMaxRowIndex = Math.Max(_scanMaxRowIndex, _rowIndex);
                            int hidden = 0;
                            double? height = null;
                            if (mode == WorksheetReadMode.Materialize)
                            {
#pragma warning disable CA1806 // Do not ignore method results
                                int.TryParse(_reader.GetAttribute(AHidden), out hidden);
#pragma warning restore CA1806 // Do not ignore method results
                                height = double.TryParse(_reader.GetAttribute(AHt), NumberStyles.Any, CultureInfo.InvariantCulture, out var ahtValue) ? Math.Abs(ahtValue) : null;
                            }

                            _pendingRowContents = true;
                            _rowOpen = true;
                            foundRow = true;
                            if (handler.OnRow(_rowIndex, hidden != 0, height) == WorksheetReadControl.Pause)
                                return WorksheetReadStatus.Paused;

                            break;
                        }
                        else if (!XmlReaderHelper.SkipContent(_reader))
                        {
                            break;
                        }
                    }

                    if (!foundRow)
                    {
                        _inSheetData = false;
                        if (mode == WorksheetReadMode.Scan && handler.OnScanResult(_scanMaxRowIndex, _scanMaxColumnIndex) == WorksheetReadControl.Pause)
                            return WorksheetReadStatus.Paused;
                    }
                }
            }

            var metadataStatus = ConsumeMetadata(ref handler);
            if (metadataStatus is { } status)
                return status;
        }
    }

    public void Dispose() => _reader.Dispose();

    private static DecodedCellValue ConvertCellValue(string rawValue, string? cellType, out CellError? error)
    {
        const NumberStyles style = NumberStyles.Any;
        error = null;
        switch (cellType)
        {
            case AS:
                if (int.TryParse(rawValue, style, CultureInfo.InvariantCulture, out var sstIndex))
                    return DecodedCellValue.SharedString(unchecked((uint)sstIndex));

                return new DecodedCellValue(rawValue, CellValueKind.FormattedString);
            case NInlineStr:
            case NStr:
                return new DecodedCellValue(Helpers.ConvertEscapeChars(rawValue), CellValueKind.FormattedString);
            case "b":
                return new DecodedCellValue(rawValue == "1");
            case "d":
                if (DateTime.TryParseExact(rawValue, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite, out var date))
                    return new DecodedCellValue(date);

                return new DecodedCellValue(rawValue, CellValueKind.FormattedString);
            case "e":
                error = ConvertError(rawValue);
                return default;
            default:
                if (double.TryParse(rawValue, style, CultureInfo.InvariantCulture, out double number))
                    return new DecodedCellValue(number);

                return new DecodedCellValue(rawValue, CellValueKind.FormattedString);
        }
    }

    private static CellError? ConvertError(string value) => value switch
    {
        "#NULL!" => CellError.NULL,
        "#DIV/0!" => CellError.DIV0,
        "#VALUE!" => CellError.VALUE,
        "#REF!" => CellError.REF,
        "#NAME?" => CellError.NAME,
        "#NUM!" => CellError.NUM,
        "#N/A" => CellError.NA,
        "#GETTING_DATA" => CellError.GETTING_DATA,
        _ => null,
    };

    private WorksheetReadStatus? ConsumeMetadata<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        if (!_metadataStarted)
        {
            _metadataStarted = true;
            if (!_reader.IsStartElement(NWorksheet, _properNamespaces.NsSpreadsheetMl) || !XmlReaderHelper.ReadFirstContent(_reader))
            {
                _endOfInput = true;
                return WorksheetReadStatus.EndOfInput;
            }
        }

        while (!_reader.EOF)
        {
            if (_reader.IsStartElement(NSheetData, _properNamespaces.NsSpreadsheetMl))
            {
                _rowIndex = -1;
                _inSheetData = XmlReaderHelper.ReadFirstContent(_reader);
                _pendingScanResult = !_inSheetData && mode == WorksheetReadMode.Scan;
                _pendingRowContents = false;
                if (handler.OnSheetDataBegin() == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
                if (_inSheetData || _pendingScanResult)
                    return null;
            }
            else if (_reader.IsStartElement(NMergeCells, _properNamespaces.NsSpreadsheetMl))
            {
                XmlReaderHelper.ReadFirstContent(_reader);
            }
            else if (_reader.IsStartElement(NMergeCell, _properNamespaces.NsSpreadsheetMl))
            {
                var cellRefs = _reader.GetAttribute(ARef);
                CellRange? range = string.IsNullOrEmpty(cellRefs) ? null : CellRange.Parse(cellRefs);
                _reader.Skip();
                if (range is { } value && handler.OnMergeCell(value) == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
            }
            else if (_reader.IsStartElement(NHeaderFooter, _properNamespaces.NsSpreadsheetMl))
            {
                var headerFooter = ReadHeaderFooter(_properNamespaces.NsSpreadsheetMl);
                if (headerFooter != null && handler.OnHeaderFooter(headerFooter) == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
            }
            else if (_reader.IsStartElement(NCols, _properNamespaces.NsSpreadsheetMl))
            {
                XmlReaderHelper.ReadFirstContent(_reader);
            }
            else if (_reader.IsStartElement(NCol, _properNamespaces.NsSpreadsheetMl))
            {
                var min = _reader.GetAttribute(AMin);
                var max = _reader.GetAttribute(AMax);
                var width = _reader.GetAttribute(AWidth);
                var customWidth = _reader.GetAttribute(ACustomWidth);
                var hidden = _reader.GetAttribute(AHidden);

                var maxVal = int.TryParse(max, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedMax) ? parsedMax : 0;
                var minVal = int.TryParse(min, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedMin) ? parsedMin : 0;
                double.TryParse(width, NumberStyles.Float, CultureInfo.InvariantCulture, out double widthVal);
                var column = new Column(minVal - 1, maxVal - 1, hidden == "1", customWidth == "1" ? widthVal : null);
                _reader.Skip();
                if (handler.OnColumn(column) == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
            }
            else if (_reader.IsStartElement(NSheetProperties, _properNamespaces.NsSpreadsheetMl))
            {
                var codeName = _reader.GetAttribute("codeName");
                _reader.Skip();
                if (handler.OnCodeName(codeName) == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
            }
            else if (_reader.IsStartElement(NSheetFormatProperties, _properNamespaces.NsSpreadsheetMl))
            {
                double? height = double.TryParse(_reader.GetAttribute(ADefaultRowHeight), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedHeight)
                    ? Math.Abs(parsedHeight)
                    : null;
                _reader.Skip();
                if (height != null && handler.OnDefaultRowHeight(height) == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
            }
            else if (_reader.IsStartElement(NDimension, _properNamespaces.NsSpreadsheetMl))
            {
                var reference = _reader.GetAttribute(ARef);
                CellRange? range = reference == null ? null : CellRange.Parse(reference);
                _reader.Skip();
                if (range is { } value && handler.OnDimension(value) == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
            }
            else if (_reader.NodeType == XmlNodeType.EndElement)
            {
                _reader.Read();
            }
            else if (!XmlReaderHelper.SkipContent(_reader))
            {
                break;
            }
        }

        _endOfInput = true;
        return WorksheetReadStatus.EndOfInput;
    }

    private HeaderFooter? ReadHeaderFooter(string nsSpreadsheetMl)
    {
        var differentFirst = _reader.GetAttribute(ADifferentFirst) == "1";
        var differentOddEven = _reader.GetAttribute(ADifferentOddEven) == "1";

        if (!XmlReaderHelper.ReadFirstContent(_reader))
            return null;

        var headerFooter = new HeaderFooter(differentFirst, differentOddEven);
        while (!_reader.EOF)
        {
            if (_reader.IsStartElement(NOddHeader, nsSpreadsheetMl))
                headerFooter.OddHeader = _reader.ReadElementContentAsString();
            else if (_reader.IsStartElement(NOddFooter, nsSpreadsheetMl))
                headerFooter.OddFooter = _reader.ReadElementContentAsString();
            else if (_reader.IsStartElement(NEvenHeader, nsSpreadsheetMl))
                headerFooter.EvenHeader = _reader.ReadElementContentAsString();
            else if (_reader.IsStartElement(NEvenFooter, nsSpreadsheetMl))
                headerFooter.EvenFooter = _reader.ReadElementContentAsString();
            else if (_reader.IsStartElement(NFirstHeader, nsSpreadsheetMl))
                headerFooter.FirstHeader = _reader.ReadElementContentAsString();
            else if (_reader.IsStartElement(NFirstFooter, nsSpreadsheetMl))
                headerFooter.FirstFooter = _reader.ReadElementContentAsString();
            else if (!XmlReaderHelper.SkipContent(_reader))
                break;
        }

        return headerFooter;
    }

    private WorksheetReadControl ConsumeCell<THandler>(int nextColumnIndex, out int columnIndex, ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        columnIndex = nextColumnIndex;
        int xfIndex = -1;
        if (ReferenceHelper.ParseReference(_reader.GetAttribute(AR), out int referenceColumn, out _))
            columnIndex = referenceColumn - 1;

        if (mode == WorksheetReadMode.Scan)
        {
            if (!_reader.IsEmptyElement)
                _scanMaxColumnIndex = Math.Max(_scanMaxColumnIndex, columnIndex);
            _reader.Skip();
            return WorksheetReadControl.Continue;
        }

        var styleAttribute = _reader.GetAttribute(AS);
        if (styleAttribute != null && int.TryParse(styleAttribute, NumberStyles.Any, CultureInfo.InvariantCulture, out int styleIndex))
            xfIndex = styleIndex;

        var cellType = _reader.GetAttribute(AT);
        if (!XmlReaderHelper.ReadFirstContent(_reader))
            return EmitCell(columnIndex, xfIndex, default, null, ref handler);

        DecodedCellValue value = default;
        CellError? error = null;
        while (!_reader.EOF)
        {
            if (_reader.NodeType == XmlNodeType.Element && _reader.LocalName == NV)
            {
                string rawValue = _reader.ReadElementContentAsString();
                if (!string.IsNullOrEmpty(rawValue))
                    value = ConvertCellValue(rawValue, cellType, out error);
            }
            else if (_reader.NodeType == XmlNodeType.Element && _reader.LocalName == NIs)
            {
                string rawValue = StringHelper.ReadStringItem(_reader, _properNamespaces.NsSpreadsheetMl);
                if (!string.IsNullOrEmpty(rawValue))
                    value = ConvertCellValue(rawValue, cellType, out error);
            }
            else if (!XmlReaderHelper.SkipContent(_reader))
            {
                break;
            }
        }

        return EmitCell(columnIndex, xfIndex, value, error, ref handler);
    }

    private WorksheetReadControl EmitCell<THandler>(int column, int style, DecodedCellValue value, CellError? error, ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        var cell = workbook.CreateCell(column, value, workbook.GetEffectiveCellStyle(style, 0), error, date1904);
        return handler.OnCell(in cell);
    }
}