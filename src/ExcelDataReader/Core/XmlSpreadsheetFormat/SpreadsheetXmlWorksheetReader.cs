#nullable enable

using System.Globalization;
using System.Text;
using System.Xml;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core.XmlSpreadsheetFormat;

internal sealed class SpreadsheetXmlWorksheetReader : IWorksheetReader
{
    private const string SpreadsheetNamespace = "urn:schemas-microsoft-com:office:spreadsheet";
    private const string ExcelNamespace = "urn:schemas-microsoft-com:office:excel";

    private readonly XmlReader _reader;
    private readonly int _worksheetIndex;
    private readonly IReadOnlyDictionary<string, ExtendedFormat> _stylesById;
    private readonly IReadOnlyDictionary<int, NumberFormatString>? _formats;
    private readonly WorksheetReadMode _mode;

    private bool _worksheetFound;
    private bool _tableFound;
    private bool _finished;
    private bool _inputEnded;
    private bool _scanResultSent;
    private bool _defaultRowHeightSent;
    private bool _dimensionSent;
    private bool _optionCodeNamePending;
    private bool _optionHeaderFooterPending;
    private int _seenWorksheetIndex = -1;
    private int _worksheetDepth = -1;
    private int _tableDepth = -1;
    private int _rowDepth = -1;
    private int _currentRowIndex;
    private int _currentRowSpan;
    private int _currentRowSpanOffset;
    private int _currentColumnIndex;
    private bool _currentRowHidden;
    private double? _currentRowHeight;
    private double _currentRowEffectiveHeight;
    private bool _rowOpen;
    private bool _rowEmpty;
    private bool _physicalRowPending;
    private int _physicalRowIndex;
    private int _physicalRowSpan;
    private bool _physicalRowHidden;
    private double? _physicalRowHeight;
    private bool _spanBeginPending;
    private bool _rowEndPending;
    private bool _tableBeginSent;
    private bool _mergePending;
    private bool _cellPending;
    private CellRange? _pendingMerge;
    private Cell _pendingCell;
    private string? _pendingCodeName;
    private HeaderFooter? _pendingHeaderFooter;
    private int _nextLogicalRow;
    private int _maxRowIndex = -1;
    private int _maxColumnIndex = -1;
    private int _columnDefinitionIndex;
    private bool _disposed;

    public SpreadsheetXmlWorksheetReader(
        Stream stream,
        int worksheetIndex,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
        IReadOnlyDictionary<int, NumberFormatString>? formats,
        WorksheetReadMode mode)
    {
        _worksheetIndex = worksheetIndex;
        _stylesById = stylesById;
        _formats = formats;
        _mode = mode;
        _reader = CreateXmlReaderAtStart(stream, tolerateLeadingWhitespace: true);
    }

    public WorksheetReadStatus Consume<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        if (_finished)
            return WorksheetReadStatus.EndOfInput;
        if (_inputEnded)
            return Finish(ref handler);

        while (true)
        {
            if (EmitPendingEvents(ref handler) == WorksheetReadControl.Pause)
                return WorksheetReadStatus.Paused;

            if (_rowOpen)
            {
                ReadCurrentRowNode();
                continue;
            }

            if (_physicalRowPending)
            {
                _physicalRowPending = false;
                if (_mode == WorksheetReadMode.Materialize && _physicalRowIndex < _nextLogicalRow)
                    throw new XmlException("SpreadsheetML rows must be in ascending index order.");

                BeginPhysicalRow();
                if (EmitRowBegin(ref handler))
                    return WorksheetReadStatus.Paused;
                continue;
            }

            if (!_worksheetFound)
            {
                if (!FindWorksheet())
                    return Finish(ref handler);
                if (_inputEnded)
                    return Finish(ref handler);
            }

            if (!_reader.Read())
                return Finish(ref handler);

            if (!_tableFound)
            {
                if (ReadWorksheetNode())
                    return Finish(ref handler);
                continue;
            }

            if (ReadTableNode(ref handler) == WorksheetReadControl.Pause)
                return WorksheetReadStatus.Paused;
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        _reader.Dispose();
    }

    internal static (string VisibleState, string? CodeName, HeaderFooter? HeaderFooter) ParseWorksheetOptions(XmlReader worksheetOptionsReader)
    {
        string visibleState = "visible";
        string? codeName = null;
        string? headerData = null;
        string? footerData = null;

        using (worksheetOptionsReader)
        {
            while (worksheetOptionsReader.Read())
            {
                ReadCurrentNode:
                if (worksheetOptionsReader.NodeType != XmlNodeType.Element ||
                    worksheetOptionsReader.NamespaceURI != ExcelNamespace)
                {
                    continue;
                }

                if (worksheetOptionsReader.LocalName == "Visible")
                {
                    var visible = worksheetOptionsReader.ReadElementContentAsString();
                    visibleState = visible switch
                    {
                        _ when string.Equals(visible, "sheethidden", StringComparison.OrdinalIgnoreCase) => "hidden",
                        _ when string.Equals(visible, "sheetveryhidden", StringComparison.OrdinalIgnoreCase) => "veryhidden",
                        _ => "visible",
                    };
                }
                else if (worksheetOptionsReader.LocalName == "CodeName")
                {
                    codeName = worksheetOptionsReader.ReadElementContentAsString();
                }
                else if (worksheetOptionsReader.LocalName == "Header")
                {
                    headerData = worksheetOptionsReader.GetAttribute("Data", ExcelNamespace);
                    worksheetOptionsReader.Skip();
                }
                else if (worksheetOptionsReader.LocalName == "Footer")
                {
                    footerData = worksheetOptionsReader.GetAttribute("Data", ExcelNamespace);
                    worksheetOptionsReader.Skip();
                }
                else
                {
                    continue;
                }

                if (worksheetOptionsReader.NodeType == XmlNodeType.Element)
                    goto ReadCurrentNode;
            }
        }

        var headerFooter = headerData != null || footerData != null
            ? new HeaderFooter(NormalizeHeaderFooter(footerData), NormalizeHeaderFooter(headerData))
            : null;
        return (visibleState, codeName, headerFooter);
    }

    internal static int ReadIndex(XmlReader reader, int current)
    {
        int index = ParseInt(GetSpreadsheetAttribute(reader, "Index"), -1);
        return index > 0 ? index - 1 : current;
    }

    internal static int ReadCount(XmlReader reader, string attribute)
        => Math.Max(0, ParseInt(GetSpreadsheetAttribute(reader, attribute)));

    private static string? NormalizeHeaderFooter(string? value)
        => value?.Replace("&V", "&L").Replace("&H", "&R").Replace("&S", "&P");

    private static bool IsCellElement(XmlReader reader)
        => reader.NodeType == XmlNodeType.Element &&
            reader.NamespaceURI == SpreadsheetNamespace &&
            reader.LocalName == "Cell";

    private static bool IsEndElementAt(XmlReader reader, int depth)
        => reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth;

    private static bool ScanCell(XmlReader cellReader)
    {
        if (cellReader.IsEmptyElement)
            return false;

        int cellDepth = cellReader.Depth;
        bool hasData = false;
        while (cellReader.Read() && !IsEndElementAt(cellReader, cellDepth))
        {
            if (cellReader.NodeType == XmlNodeType.Element &&
                cellReader.Depth == cellDepth + 1 &&
                cellReader.NamespaceURI == SpreadsheetNamespace &&
                cellReader.LocalName == "Data")
            {
                hasData = true;
            }
        }

        return hasData;
    }

    private static bool IsDurationStyle(ExtendedFormat style, IReadOnlyDictionary<int, NumberFormatString>? formats)
    {
        if (style.NumberFormatIndex != 0)
        {
            NumberFormatString? fmt = null;
            formats?.TryGetValue(style.NumberFormatIndex, out fmt);
#pragma warning disable CA1305
            fmt ??= BuiltinNumberFormat.GetBuiltinNumberFormat(style.NumberFormatIndex);
#pragma warning restore CA1305
            return fmt?.IsTimeSpanFormat == true;
        }

        return false;
    }

    private static object? ParseCellValue(XmlReader cellReader, ExtendedFormat style, IReadOnlyDictionary<int, NumberFormatString>? formats, out CellError? cellError, out bool hasData)
    {
        cellError = null;
        hasData = false;
        if (cellReader.IsEmptyElement)
            return null;

        int cellDepth = cellReader.Depth;
        object? result = null;
        while (cellReader.Read() && !IsEndElementAt(cellReader, cellDepth))
        {
            if (cellReader.NodeType != XmlNodeType.Element ||
                cellReader.Depth != cellDepth + 1 ||
                cellReader.NamespaceURI != SpreadsheetNamespace ||
                cellReader.LocalName != "Data")
            {
                continue;
            }

            hasData = true;
            var type = GetSpreadsheetAttribute(cellReader, "Type") ?? string.Empty;
            var text = ReadElementText(cellReader);
            result = ParseDataText(type, text, IsDurationStyle(style, formats), out cellError);
            while (!IsEndElementAt(cellReader, cellDepth) && cellReader.Read())
            {
            }

            break;
        }

        return result;
    }

    private static object? ParseDataText(string type, string value, bool duration, out CellError? cellError)
    {
        cellError = null;
        switch (type)
        {
            case "Number":
                double number = ParseDouble(value, 0D);
                return duration ? TimeSpan.FromDays(number) : number;
            case "DateTime":
                if (TryParseSpreadsheetDateTime(value, duration, out var dateTime))
                    return dateTime;
                break;
            case "Boolean":
                return Helpers.BoxBoolean(ParseBool(value));
            case "Error":
                cellError = ParseCellError(value);
                return null;
        }

        return duration && TimeSpan.TryParse(value, CultureInfo.InvariantCulture, out var timeSpan) ? timeSpan : value;
    }

    private static CellError? ParseCellError(string value) => value switch
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

    private static ExtendedFormat GetDefaultStyle(IReadOnlyDictionary<string, ExtendedFormat> stylesById)
        => stylesById.TryGetValue("Default", out var defaultStyle) ? defaultStyle : ExtendedFormat.Zero;

    private static double? ParseColumnWidth(string? value)
    {
        var width = ParseNullableDouble(value);
        return width / 5.25D;
    }

    private static string ReadElementText(XmlReader reader)
    {
        if (reader.IsEmptyElement)
        {
            reader.Read();
            return string.Empty;
        }

        string? result = null;
        StringBuilder? builder = null;
        int depth = reader.Depth;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)
                break;
            if (reader.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.SignificantWhitespace or XmlNodeType.Whitespace)
            {
                if (result == null)
                    result = reader.Value;
                else
                    (builder ??= new StringBuilder(result)).Append(reader.Value);
            }
        }

        return builder?.ToString() ?? result ?? string.Empty;
    }

    private static bool TryParseSpreadsheetDateTime(string value, bool duration, out object? dateTime)
    {
        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsed))
        {
            dateTime = duration ? TimeSpan.FromDays(GetExcelSerialDate(parsed)) : parsed;
            return true;
        }

        const string fakeLeapDay = "1900-02-29";
        if (value.StartsWith(fakeLeapDay, StringComparison.Ordinal) &&
            TimeSpan.TryParse(value.Substring(fakeLeapDay.Length).TrimStart('T'), CultureInfo.InvariantCulture, out var time))
        {
            // Validate the DateTime even when returning a duration, as on the ordinary date path.
            var date = new DateTime(1900, 2, 28).Add(time);
            dateTime = duration ? TimeSpan.FromDays(60D + time.TotalDays) : date;
            return true;
        }

        dateTime = default;
        return false;
    }

    private static double GetExcelSerialDate(DateTime dateTime)
    {
        var serialDate = (dateTime - new DateTime(1899, 12, 31)).TotalDays;
        if (dateTime >= new DateTime(1900, 3, 1))
            serialDate++;
        return serialDate;
    }

    private static XmlReader CreateXmlReaderAtStart(Stream stream, bool tolerateLeadingWhitespace)
    {
        if (stream.CanSeek)
        {
            stream.Seek(0, SeekOrigin.Begin);
            if (tolerateLeadingWhitespace)
                SkipLeadingAsciiWhitespace(stream);
        }

        return XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreWhitespace = false,
            CloseInput = false,
        });
    }

    private static void SkipLeadingAsciiWhitespace(Stream stream)
    {
        while (true)
        {
            int value = stream.ReadByte();
            if (value < 0)
                return;
            if (!IsAsciiWhitespace((byte)value))
            {
                stream.Seek(-1, SeekOrigin.Current);
                return;
            }
        }
    }

    private static bool IsAsciiWhitespace(byte value)
        => value == (byte)' ' || value == (byte)'\t' || value == (byte)'\r' || value == (byte)'\n';

    private static string? GetSpreadsheetAttribute(XmlReader reader, string attributeName)
        => reader.GetAttribute(attributeName, SpreadsheetNamespace) ?? reader.GetAttribute(attributeName);

    private static int ParseInt(string? value, int defaultValue = 0)
        => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : defaultValue;

    private static double ParseDouble(string? value, double defaultValue)
        => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result) ? result : defaultValue;

    private static double? ParseNullableDouble(string? value)
        => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result) ? result : null;

    private static bool ParseBool(string? value)
        => value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);

    private WorksheetReadControl EmitPendingEvents<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        if (!_defaultRowHeightSent)
        {
            _defaultRowHeightSent = true;
            if (handler.OnDefaultRowHeight(15D) == WorksheetReadControl.Pause)
                return WorksheetReadControl.Pause;
        }

        while (true)
        {
            if (_optionCodeNamePending)
            {
                _optionCodeNamePending = false;
                if (handler.OnCodeName(_pendingCodeName) == WorksheetReadControl.Pause)
                    return WorksheetReadControl.Pause;
                continue;
            }

            if (_optionHeaderFooterPending)
            {
                _optionHeaderFooterPending = false;
                if (_pendingHeaderFooter != null && handler.OnHeaderFooter(_pendingHeaderFooter) == WorksheetReadControl.Pause)
                    return WorksheetReadControl.Pause;
                continue;
            }

            if (_mergePending)
            {
                _mergePending = false;
                if (handler.OnMergeCell(_pendingMerge!) == WorksheetReadControl.Pause)
                    return WorksheetReadControl.Pause;
                continue;
            }

            if (_cellPending)
            {
                _cellPending = false;
                if (handler.OnCell(in _pendingCell) == WorksheetReadControl.Pause)
                    return WorksheetReadControl.Pause;
                continue;
            }

            if (_spanBeginPending)
            {
                _spanBeginPending = false;
                _currentRowIndex++;
                _currentRowSpanOffset++;
                _currentRowHeight = _currentRowEffectiveHeight;
                _rowOpen = true;
                _rowEmpty = true;
                if (EmitRowBegin(ref handler))
                    return WorksheetReadControl.Pause;
                continue;
            }

            if (_rowEndPending)
            {
                _rowEndPending = false;
                bool hasSpannedRows = _mode == WorksheetReadMode.Materialize &&
                    _currentRowSpanOffset < _currentRowSpan;
                if (hasSpannedRows)
                {
                    _spanBeginPending = true;
                }
                else
                {
                    _rowOpen = false;
                    _nextLogicalRow = _mode == WorksheetReadMode.Scan
                        ? _currentRowIndex + _currentRowSpan + 1
                        : _currentRowIndex + 1;
                }

                if (_mode == WorksheetReadMode.Materialize && handler.OnRowEnd() == WorksheetReadControl.Pause)
                    return WorksheetReadControl.Pause;
                continue;
            }

            return WorksheetReadControl.Continue;
        }
    }

    private void ReadCurrentRowNode()
    {
        if (_rowEmpty)
        {
            _rowEmpty = false;
            _rowEndPending = true;
            return;
        }

        if (!_reader.Read())
        {
            _rowOpen = false;
            return;
        }

        if (_reader.NodeType == XmlNodeType.EndElement && _reader.Depth == _rowDepth)
        {
            _rowEndPending = true;
            return;
        }

        if (!IsCellElement(_reader))
            return;

        _currentColumnIndex = ReadIndex(_reader, _currentColumnIndex);
        int mergeAcross = ReadCount(_reader, "MergeAcross");
        int mergeDown = ReadCount(_reader, "MergeDown");
        var styleId = GetSpreadsheetAttribute(_reader, "StyleID");
        var style = styleId != null && _stylesById.TryGetValue(styleId, out var foundStyle)
            ? foundStyle
            : GetDefaultStyle(_stylesById);

        object? value = null;
        CellError? error = null;
        bool hasData;
        if (_mode == WorksheetReadMode.Materialize)
        {
            value = ParseCellValue(_reader, style, _formats, out error, out hasData);
        }
        else
        {
            hasData = ScanCell(_reader);
        }

        if (hasData || error != null || mergeAcross > 0 || mergeDown > 0)
        {
            _maxColumnIndex = Math.Max(_maxColumnIndex, _currentColumnIndex + mergeAcross);
            if (mergeAcross > 0 || mergeDown > 0)
            {
                _pendingMerge = new CellRange(
                    _currentColumnIndex,
                    _currentRowIndex,
                    _currentColumnIndex + mergeAcross,
                    _currentRowIndex + mergeDown);
                _mergePending = true;
            }

            if (_mode == WorksheetReadMode.Materialize)
            {
                _pendingCell = new Cell(_currentColumnIndex, value, style, error);
                _cellPending = true;
            }
        }

        _currentColumnIndex += mergeAcross + 1;
    }

    private bool ReadWorksheetNode()
    {
        if (_reader.NodeType == XmlNodeType.Element &&
            _reader.Depth == _worksheetDepth + 1 &&
            _reader.LocalName == "WorksheetOptions" &&
            _reader.NamespaceURI == ExcelNamespace)
        {
            var (_, codeName, headerFooter) = ParseWorksheetOptions(_reader.ReadSubtree());
            _pendingCodeName = codeName;
            _pendingHeaderFooter = headerFooter;
            _optionCodeNamePending = true;
            _optionHeaderFooterPending = headerFooter != null;
        }
        else if (_reader.NodeType == XmlNodeType.Element &&
            _reader.Depth == _worksheetDepth + 1 &&
            _reader.LocalName == "Table" &&
            _reader.NamespaceURI == SpreadsheetNamespace)
        {
            _tableFound = !_reader.IsEmptyElement;
            _tableDepth = _reader.Depth;
        }

        return _reader.NodeType == XmlNodeType.EndElement && _reader.Depth == _worksheetDepth;
    }

    private WorksheetReadControl ReadTableNode<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        if (_reader.NodeType == XmlNodeType.EndElement && _reader.Depth == _tableDepth)
        {
            _tableFound = false;
            return WorksheetReadControl.Continue;
        }

        if (_reader.NodeType != XmlNodeType.Element || _reader.Depth != _tableDepth + 1 ||
            _reader.NamespaceURI != SpreadsheetNamespace)
        {
            return WorksheetReadControl.Continue;
        }

        if (_reader.LocalName == "Column")
        {
            _columnDefinitionIndex = ReadIndex(_reader, _columnDefinitionIndex);
            int span = ReadCount(_reader, "Span");
            bool hidden = ParseBool(GetSpreadsheetAttribute(_reader, "Hidden"));
            double? width = ParseColumnWidth(GetSpreadsheetAttribute(_reader, "Width"));
            var column = new Column(_columnDefinitionIndex, _columnDefinitionIndex + span, hidden, width);
            _columnDefinitionIndex += span + 1;
            return handler.OnColumn(column);
        }

        if (_reader.LocalName == "Row")
        {
            _physicalRowIndex = ReadIndex(_reader, _nextLogicalRow);
            _physicalRowSpan = ReadCount(_reader, "Span");
            _physicalRowHidden = ParseBool(GetSpreadsheetAttribute(_reader, "Hidden"));
            _physicalRowHeight = ParseNullableDouble(GetSpreadsheetAttribute(_reader, "Height"));
            _physicalRowPending = true;
            if (!_tableBeginSent)
            {
                _tableBeginSent = true;
                return handler.OnSheetDataBegin();
            }
        }

        return WorksheetReadControl.Continue;
    }

    private bool FindWorksheet()
    {
        while (_reader.Read())
        {
            if (_reader.NodeType == XmlNodeType.Element &&
                _reader.LocalName == "Worksheet" &&
                _reader.NamespaceURI == SpreadsheetNamespace)
            {
                _seenWorksheetIndex++;
                if (_seenWorksheetIndex == _worksheetIndex)
                {
                    _worksheetFound = true;
                    _worksheetDepth = _reader.Depth;
                    _inputEnded = _reader.IsEmptyElement;
                    return true;
                }
            }
        }

        return false;
    }

    private void BeginPhysicalRow()
    {
        _currentRowIndex = _physicalRowIndex;
        _currentRowSpan = _physicalRowSpan;
        _currentRowSpanOffset = 0;
        _currentRowHidden = _physicalRowHidden;
        _currentRowHeight = _physicalRowHeight;
        _currentRowEffectiveHeight = _physicalRowHidden ? 0D : _physicalRowHeight ?? 15D;
        _currentColumnIndex = 0;
        _rowDepth = _reader.Depth;
        _rowEmpty = _reader.IsEmptyElement;
        _rowOpen = true;
        _maxRowIndex = Math.Max(_maxRowIndex, _currentRowIndex + _currentRowSpan);
    }

    private bool EmitRowBegin<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        if (_mode != WorksheetReadMode.Materialize)
            return false;

        return handler.OnRow(_currentRowIndex, _currentRowHidden, _currentRowHeight) == WorksheetReadControl.Pause;
    }

    private WorksheetReadStatus Finish<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        _inputEnded = true;
        if (!_scanResultSent)
        {
            if (!_dimensionSent)
            {
                _dimensionSent = true;
                if (_maxRowIndex >= 0 && _maxColumnIndex >= 0 &&
                    handler.OnDimension(new CellRange(0, 0, _maxColumnIndex, _maxRowIndex)) == WorksheetReadControl.Pause)
                {
                    return WorksheetReadStatus.Paused;
                }
            }

            _scanResultSent = true;
            if (handler.OnScanResult(_maxRowIndex, _maxColumnIndex) == WorksheetReadControl.Pause)
                return WorksheetReadStatus.Paused;
        }

        _finished = true;
        return WorksheetReadStatus.EndOfInput;
    }
}