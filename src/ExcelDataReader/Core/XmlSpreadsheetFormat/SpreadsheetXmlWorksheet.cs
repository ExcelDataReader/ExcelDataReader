#nullable enable

using System.Globalization;
using System.Text;
using System.Xml;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core.XmlSpreadsheetFormat;

internal sealed class SpreadsheetXmlWorksheet : IWorksheet
{
    private const string SpreadsheetNamespace = "urn:schemas-microsoft-com:office:spreadsheet";
    private const string ExcelNamespace = "urn:schemas-microsoft-com:office:excel";

    private readonly List<Row> _rows;
    private readonly Stream? _stream;
    private readonly IReadOnlyDictionary<string, ExtendedFormat>? _stylesById;
    private readonly IReadOnlyDictionary<int, NumberFormatString>? _formats;
    private readonly int _worksheetIndex;

#pragma warning disable CS8618 // CodeName is intentionally nullable (null for sheets without x:CodeName), matching CsvWorksheet pattern
    private SpreadsheetXmlWorksheet(
        string name,
        string visibleState,
        string? codeName,
        HeaderFooter? headerFooter,
        List<Row> rows,
        List<Column> columnWidths,
        CellRange[] mergeCells,
        int fieldCount,
        int rowCount,
        CellRange? dimension,
        Stream? stream = null,
        IReadOnlyDictionary<string, ExtendedFormat>? stylesById = null,
        IReadOnlyDictionary<int, NumberFormatString>? formats = null,
        int worksheetIndex = -1)
    {
        Name = name;
#pragma warning disable CS8601 // Intentional: CodeName can be null for sheets without x:CodeName, matching the pattern of CsvWorksheet.CodeName
        CodeName = codeName;
#pragma warning restore CS8601
        VisibleState = visibleState;
        HeaderFooter = headerFooter;
        _rows = rows;
        ColumnWidths = columnWidths;
        MergeCells = mergeCells;
        FieldCount = fieldCount;
        RowCount = rowCount;
        Dimension = dimension;
        _stream = stream;
        _stylesById = stylesById;
        _formats = formats;
        _worksheetIndex = worksheetIndex;
    }
#pragma warning restore CS8618

    public string Name { get; }

    public string CodeName { get; }

    public string VisibleState { get; }

    public HeaderFooter? HeaderFooter { get; }

    public int FieldCount { get; }

    public int RowCount { get; }

    public CellRange? Dimension { get; }

    public CellRange[] MergeCells { get; }

    public List<Column> ColumnWidths { get; }

    public static SpreadsheetXmlWorksheet Create(
        Stream stream,
        int worksheetIndex,
        string name,
        string visibleState,
        string? codeName,
        HeaderFooter? headerFooter,
        int expandedColumnCount,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
        IReadOnlyDictionary<int, NumberFormatString>? formats,
        bool singlePassMode)
    {
        if (singlePassMode)
        {
            return new SpreadsheetXmlWorksheet(
                name,
                visibleState,
                codeName,
                headerFooter,
                [],
                [],
                [],
                0,
                0,
                null,
                stream,
                stylesById,
                formats,
                worksheetIndex);
        }

        using var workbookReader = CreateXmlReaderAtStart(stream, tolerateLeadingWhitespace: true);
        int currentWorksheetIndex = -1;
        while (workbookReader.Read())
        {
            if (workbookReader.NodeType != XmlNodeType.Element ||
                workbookReader.LocalName != "Worksheet" ||
                workbookReader.NamespaceURI != SpreadsheetNamespace)
            {
                continue;
            }

            currentWorksheetIndex++;
            if (currentWorksheetIndex == worksheetIndex)
            {
                return Parse(workbookReader.ReadSubtree(), stylesById, formats);
            }
        }

        throw new XmlException("Invalid SpreadsheetML worksheet.");
    }

    public static SpreadsheetXmlWorksheet Parse(XmlReader worksheetReader, IReadOnlyDictionary<string, ExtendedFormat> stylesById, IReadOnlyDictionary<int, NumberFormatString>? formats = null)
    {
        using (worksheetReader)
        {
            if (!worksheetReader.Read() || worksheetReader.NodeType != XmlNodeType.Element)
                throw new XmlException("Invalid SpreadsheetML worksheet.");

            var name = GetSpreadsheetAttribute(worksheetReader, "Name") ?? string.Empty;
            var visibleState = "visible";
            string? codeName = null;
            HeaderFooter? headerFooter = null;
            var rows = new List<Row>();
            var columnWidths = new List<Column>();
            var mergeCells = new List<CellRange>();
            int maxColumn = -1;
            int maxRow = -1;

            while (worksheetReader.Read())
            {
                if (worksheetReader.NodeType != XmlNodeType.Element)
                    continue;

                if (worksheetReader.LocalName == "WorksheetOptions" && worksheetReader.NamespaceURI == ExcelNamespace)
                {
                    (visibleState, codeName, headerFooter) = ParseWorksheetOptions(worksheetReader.ReadSubtree());
                }
                else if (worksheetReader.LocalName == "Table" && worksheetReader.NamespaceURI == SpreadsheetNamespace)
                {
                    ParseTable(
                        worksheetReader.ReadSubtree(),
                        stylesById,
                        formats,
                        rows,
                        columnWidths,
                        mergeCells,
                        ref maxColumn,
                        ref maxRow);
                }
            }

            int fieldCount = maxColumn + 1;
            int rowCount = maxRow + 1;

            var normalizedRows = NormalizeRows(rows, rowCount);
            var dimension = fieldCount > 0 && rowCount > 0
                ? new CellRange(0, 0, fieldCount - 1, rowCount - 1)
                : null;

            return new SpreadsheetXmlWorksheet(
                name,
                visibleState,
                codeName,
                headerFooter,
                normalizedRows,
                columnWidths,
                [.. mergeCells],
                fieldCount,
                rowCount,
                dimension);
        }
    }

    public IEnumerable<Row> ReadRows()
    {
        if (_stream != null && _stylesById != null)
        {
            foreach (var row in StreamRows(_stream, _worksheetIndex, _stylesById, _formats))
                yield return row;
            yield break;
        }

        foreach (var row in _rows)
            yield return row;
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
                    visibleState = visible.ToLowerInvariant() switch
                    {
                        "sheethidden" => "hidden",
                        "sheetveryhidden" => "veryhidden",
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

                // ReadElementContentAsString/Skip already advanced past the end tag.
                // If the reader is now at another element, handle it without calling Read() first.
                if (worksheetOptionsReader.NodeType == XmlNodeType.Element)
                    goto ReadCurrentNode;
            }
        }

        var headerFooter = headerData != null || footerData != null
            ? new HeaderFooter(NormalizeHeaderFooter(footerData), NormalizeHeaderFooter(headerData))
            : null;

        return (visibleState, codeName, headerFooter);
    }

    // Keep for callers that only need the visible state (e.g. workbook-level scan).
    internal static string ParseVisibleState(XmlReader worksheetOptionsReader)
        => ParseWorksheetOptions(worksheetOptionsReader).VisibleState;

    private static IEnumerable<Row> StreamRows(
        Stream stream,
        int worksheetIndex,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
        IReadOnlyDictionary<int, NumberFormatString>? formats)
    {
        using var workbookReader = CreateXmlReaderAtStart(stream, tolerateLeadingWhitespace: true);
        int currentWorksheetIndex = -1;
        while (workbookReader.Read())
        {
            if (workbookReader.NodeType != XmlNodeType.Element ||
                workbookReader.LocalName != "Worksheet" ||
                workbookReader.NamespaceURI != SpreadsheetNamespace)
            {
                continue;
            }

            currentWorksheetIndex++;
            if (currentWorksheetIndex != worksheetIndex)
                continue;

            foreach (var row in StreamRowsFromWorksheet(workbookReader.ReadSubtree(), stylesById, formats))
                yield return row;
            yield break;
        }
    }

    private static IEnumerable<Row> StreamRowsFromWorksheet(
        XmlReader worksheetReader,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
        IReadOnlyDictionary<int, NumberFormatString>? formats)
    {
        using (worksheetReader)
        {
            if (!worksheetReader.Read() || worksheetReader.NodeType != XmlNodeType.Element)
                throw new XmlException("Invalid SpreadsheetML worksheet.");

            while (worksheetReader.Read())
            {
                if (worksheetReader.NodeType == XmlNodeType.Element &&
                    worksheetReader.LocalName == "Table" &&
                    worksheetReader.NamespaceURI == SpreadsheetNamespace)
                {
                    foreach (var row in StreamRowsFromTable(worksheetReader.ReadSubtree(), stylesById, formats))
                        yield return row;
                    yield break;
                }
            }
        }
    }

    private static IEnumerable<Row> StreamRowsFromTable(
        XmlReader tableReader,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
        IReadOnlyDictionary<int, NumberFormatString>? formats)
    {
        using (tableReader)
        {
            if (!tableReader.Read() || tableReader.NodeType != XmlNodeType.Element)
                yield break;

            int currentRowIndex = 0;
            int maxColumn = -1;
            int maxRow = -1;
            var mergeCells = new List<CellRange>();

            while (tableReader.Read())
            {
                if (tableReader.NodeType != XmlNodeType.Element ||
                    tableReader.NamespaceURI != SpreadsheetNamespace ||
                    tableReader.LocalName != "Row")
                {
                    continue;
                }

                var row = ParseRowForStreaming(
                    tableReader.ReadSubtree(),
                    stylesById,
                    formats,
                    mergeCells,
                    ref currentRowIndex,
                    ref maxColumn,
                    ref maxRow,
                    out var rowSpan);

                while (row.RowIndex > currentRowIndex)
                {
                    yield return new Row(currentRowIndex, 15D, []);
                    currentRowIndex++;
                }

                yield return row;
                for (int i = 1; i <= rowSpan; i++)
                {
                    yield return new Row(row.RowIndex + i, row.Height, []);
                }

                currentRowIndex = row.RowIndex + rowSpan + 1;
            }
        }
    }

    private static Row ParseRowForStreaming(
        XmlReader rowReader,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
        IReadOnlyDictionary<int, NumberFormatString>? formats,
        List<CellRange> mergeCells,
        ref int currentRowIndex,
        ref int maxColumn,
        ref int maxRow,
        out int rowSpan)
    {
        using (rowReader)
        {
            rowSpan = 0;
            if (!rowReader.Read() || rowReader.NodeType != XmlNodeType.Element)
                return new Row(currentRowIndex, 15D, []);

            int rowIndexFromAttribute = ParseInt(GetSpreadsheetAttribute(rowReader, "Index"), -1);
            if (rowIndexFromAttribute > 0)
                currentRowIndex = rowIndexFromAttribute - 1;

            bool hidden = ParseBool(GetSpreadsheetAttribute(rowReader, "Hidden"));
            rowSpan = Math.Max(0, ParseInt(GetSpreadsheetAttribute(rowReader, "Span")));
            double rowHeight = hidden ? 0D : ParseDouble(GetSpreadsheetAttribute(rowReader, "Height"), 15D);
            var cells = new List<Cell>();
            int currentColumnIndex = 0;

            while (rowReader.Read())
            {
                if (rowReader.NodeType != XmlNodeType.Element ||
                    rowReader.NamespaceURI != SpreadsheetNamespace ||
                    rowReader.LocalName != "Cell")
                {
                    continue;
                }

                int columnIndexFromAttribute = ParseInt(GetSpreadsheetAttribute(rowReader, "Index"), -1);
                if (columnIndexFromAttribute > 0)
                    currentColumnIndex = columnIndexFromAttribute - 1;

                int mergeAcross = Math.Max(0, ParseInt(GetSpreadsheetAttribute(rowReader, "MergeAcross")));
                int mergeDown = Math.Max(0, ParseInt(GetSpreadsheetAttribute(rowReader, "MergeDown")));
                var styleId = GetSpreadsheetAttribute(rowReader, "StyleID");
                var effectiveStyle = styleId != null && stylesById.TryGetValue(styleId, out var style)
                    ? style
                    : GetDefaultStyle(stylesById);

                bool hasData = false;
                var rawValue = ParseCellValue(rowReader.ReadSubtree(), out var cellError, out hasData);
                var value = ConvertCellValue(rawValue, effectiveStyle, formats);

                if (hasData || cellError != null || mergeAcross > 0 || mergeDown > 0)
                {
                    cells.Add(new Cell(currentColumnIndex, value, effectiveStyle, cellError));
                    maxColumn = Math.Max(maxColumn, currentColumnIndex + mergeAcross);
                }

                if (mergeAcross > 0 || mergeDown > 0)
                    mergeCells.Add(new CellRange(currentColumnIndex, currentRowIndex, currentColumnIndex + mergeAcross, currentRowIndex + mergeDown));

                currentColumnIndex += mergeAcross + 1;
            }

            maxRow = Math.Max(maxRow, currentRowIndex);
            return new Row(currentRowIndex, rowHeight, cells);
        }
    }

    private static void ParseTable(
        XmlReader tableReader,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
        IReadOnlyDictionary<int, NumberFormatString>? formats,
        List<Row> rows,
        List<Column> columnWidths,
        List<CellRange> mergeCells,
        ref int maxColumn,
        ref int maxRow)
    {
        using (tableReader)
        {
            if (!tableReader.Read() || tableReader.NodeType != XmlNodeType.Element)
                return;

            int currentRowIndex = 0;
            int currentColumnDefinition = 0;

            while (tableReader.Read())
            {
                if (tableReader.NodeType != XmlNodeType.Element || tableReader.NamespaceURI != SpreadsheetNamespace)
                    continue;

                if (tableReader.LocalName == "Column")
                {
                    ParseColumn(tableReader, columnWidths, ref currentColumnDefinition);
                }
                else if (tableReader.LocalName == "Row")
                {
                    ParseRow(
                        tableReader.ReadSubtree(),
                        stylesById,
                        formats,
                        mergeCells,
                        rows,
                        ref currentRowIndex,
                        ref maxColumn,
                        ref maxRow);
                }
            }
        }
    }

    private static void ParseColumn(XmlReader columnReader, List<Column> columnWidths, ref int currentColumn)
    {
        int columnIndexFromAttribute = ParseInt(GetSpreadsheetAttribute(columnReader, "Index"), -1);
        if (columnIndexFromAttribute > 0)
        {
            currentColumn = columnIndexFromAttribute - 1;
        }

        int span = Math.Max(0, ParseInt(GetSpreadsheetAttribute(columnReader, "Span")));
        bool hidden = ParseBool(GetSpreadsheetAttribute(columnReader, "Hidden"));
        double? width = ParseColumnWidth(GetSpreadsheetAttribute(columnReader, "Width"));
        columnWidths.Add(new Column(currentColumn, currentColumn + span, hidden, width));
        currentColumn += span + 1;
    }

    private static void ParseRow(
        XmlReader rowReader,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
        IReadOnlyDictionary<int, NumberFormatString>? formats,
        List<CellRange> mergeCells,
        List<Row> rows,
        ref int currentRowIndex,
        ref int maxColumn,
        ref int maxRow)
    {
        using (rowReader)
        {
            if (!rowReader.Read() || rowReader.NodeType != XmlNodeType.Element)
                return;

            int rowIndexFromAttribute = ParseInt(GetSpreadsheetAttribute(rowReader, "Index"), -1);
            if (rowIndexFromAttribute > 0)
            {
                currentRowIndex = rowIndexFromAttribute - 1;
            }

            bool hidden = ParseBool(GetSpreadsheetAttribute(rowReader, "Hidden"));
            int rowSpan = Math.Max(0, ParseInt(GetSpreadsheetAttribute(rowReader, "Span")));
            double rowHeight = hidden ? 0D : ParseDouble(GetSpreadsheetAttribute(rowReader, "Height"), 15D);
            var cells = new List<Cell>();
            int currentColumnIndex = 0;

            while (rowReader.Read())
            {
                if (rowReader.NodeType != XmlNodeType.Element ||
                    rowReader.NamespaceURI != SpreadsheetNamespace ||
                    rowReader.LocalName != "Cell")
                {
                    continue;
                }

                int columnIndexFromAttribute = ParseInt(GetSpreadsheetAttribute(rowReader, "Index"), -1);
                if (columnIndexFromAttribute > 0)
                {
                    currentColumnIndex = columnIndexFromAttribute - 1;
                }

                int mergeAcross = Math.Max(0, ParseInt(GetSpreadsheetAttribute(rowReader, "MergeAcross")));
                int mergeDown = Math.Max(0, ParseInt(GetSpreadsheetAttribute(rowReader, "MergeDown")));
                var styleId = GetSpreadsheetAttribute(rowReader, "StyleID");
                var effectiveStyle = styleId != null && stylesById.TryGetValue(styleId, out var style)
                    ? style
                    : GetDefaultStyle(stylesById);

                bool hasData = false;
                var rawValue = ParseCellValue(rowReader.ReadSubtree(), out var cellError, out hasData);
                var value = ConvertCellValue(rawValue, effectiveStyle, formats);

                if (hasData || cellError != null || mergeAcross > 0 || mergeDown > 0)
                {
                    cells.Add(new Cell(currentColumnIndex, value, effectiveStyle, cellError));
                    maxColumn = Math.Max(maxColumn, currentColumnIndex + mergeAcross);
                }

                if (mergeAcross > 0 || mergeDown > 0)
                {
                    mergeCells.Add(new CellRange(currentColumnIndex, currentRowIndex, currentColumnIndex + mergeAcross, currentRowIndex + mergeDown));
                }

                currentColumnIndex += mergeAcross + 1;
            }

            rows.Add(new Row(currentRowIndex, rowHeight, cells));
            for (int i = 1; i <= rowSpan; i++)
            {
                rows.Add(new Row(currentRowIndex + i, rowHeight, []));
            }

            maxRow = Math.Max(maxRow, currentRowIndex + rowSpan);
            currentRowIndex += rowSpan + 1;
        }
    }

    private static object? ConvertCellValue(object? value, ExtendedFormat style, IReadOnlyDictionary<int, NumberFormatString>? formats)
    {
        if (style.NumberFormatIndex != 0)
        {
            NumberFormatString? fmt = null;
            formats?.TryGetValue(style.NumberFormatIndex, out fmt);
#pragma warning disable CA1305 // Intentional: locale-independent check for IsTimeSpanFormat
            fmt ??= BuiltinNumberFormat.GetBuiltinNumberFormat(style.NumberFormatIndex);
#pragma warning restore CA1305
            if (fmt?.IsTimeSpanFormat == true)
            {
                if (value is double number)
                    return TimeSpan.FromDays(number);

                if (value is SpreadsheetXmlDateTime dateTime)
                    return TimeSpan.FromDays(dateTime.SerialDate);

                if (value is string text && TimeSpan.TryParse(text, CultureInfo.InvariantCulture, out var timeSpan))
                    return timeSpan;
            }
        }

        if (value is SpreadsheetXmlDateTime date)
            return date.Value;

        return value;
    }

    private static object? ParseCellValue(XmlReader cellReader, out CellError? cellError, out bool hasData)
    {
        cellError = null;
        hasData = false;

        using (cellReader)
        {
            while (cellReader.Read())
            {
                if (cellReader.NodeType != XmlNodeType.Element ||
                    cellReader.NamespaceURI != SpreadsheetNamespace ||
                    cellReader.LocalName != "Data")
                {
                    continue;
                }

                hasData = true;
                var type = GetSpreadsheetAttribute(cellReader, "Type") ?? string.Empty;
                var text = ReadElementText(cellReader);

                return ParseDataText(type, text, out cellError);
            }
        }

        return null;
    }

    private static object? ParseDataText(string type, string value, out CellError? cellError)
    {
        cellError = null;
        switch (type)
        {
            case "Number":
                return ParseDouble(value, 0D);
            case "DateTime":
                if (TryParseSpreadsheetDateTime(value, out var dateTime))
                    return dateTime;
                return value;
            case "Boolean":
                return ParseBool(value);
            case "Error":
                cellError = ParseCellError(value);
                return null;
            case "String":
            default:
                return value;
        }
    }

    private static List<Row> NormalizeRows(List<Row> rows, int rowCount)
    {
        if (rows.Count >= rowCount)
            return rows;

        var rowsByIndex = rows.ToDictionary(r => r.RowIndex);
        var normalized = new List<Row>(rowCount);
        for (int i = 0; i < rowCount; i++)
        {
            if (rowsByIndex.TryGetValue(i, out var row))
            {
                normalized.Add(row);
            }
            else
            {
                normalized.Add(new Row(i, 15D, []));
            }
        }

        return normalized;
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

    private static string? NormalizeHeaderFooter(string? value)
        => value?
            .Replace("&V", "&L")
            .Replace("&H", "&R")
            .Replace("&S", "&P");

    private static ExtendedFormat GetDefaultStyle(IReadOnlyDictionary<string, ExtendedFormat> stylesById)
        => stylesById.TryGetValue("Default", out var defaultStyle) ? defaultStyle : ExtendedFormat.Zero;

    private static double? ParseColumnWidth(string? value)
    {
        var width = ParseNullableDouble(value);
        if (width == null)
            return null;

        // SpreadsheetML stores column widths in points. The reader exposes the
        // Excel character-width unit used by the binary and OpenXml readers.
        return width / 5.25D;
    }

    private static string ReadElementText(XmlReader reader)
    {
        if (reader.IsEmptyElement)
        {
            reader.Read();
            return string.Empty;
        }

        var result = new StringBuilder();
        var depth = reader.Depth;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)
                break;

            if (reader.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.SignificantWhitespace or XmlNodeType.Whitespace)
            {
                result.Append(reader.Value);
            }
        }

        return result.ToString();
    }

    private static bool TryParseSpreadsheetDateTime(string value, out SpreadsheetXmlDateTime dateTime)
    {
        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsed))
        {
            dateTime = new SpreadsheetXmlDateTime(parsed, GetExcelSerialDate(parsed));
            return true;
        }

        const string fakeLeapDay = "1900-02-29";
        if (value.StartsWith(fakeLeapDay, StringComparison.Ordinal) &&
            TimeSpan.TryParse(value.Substring(fakeLeapDay.Length).TrimStart('T'), CultureInfo.InvariantCulture, out var time))
        {
            dateTime = new SpreadsheetXmlDateTime(new DateTime(1900, 2, 28).Add(time), 60D + time.TotalDays);
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

        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreWhitespace = false,
            CloseInput = false,
        };

        return XmlReader.Create(stream, settings);
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

    private readonly struct SpreadsheetXmlDateTime(DateTime value, double serialDate)
    {
        public DateTime Value { get; } = value;

        public double SerialDate { get; } = serialDate;
    }
}
