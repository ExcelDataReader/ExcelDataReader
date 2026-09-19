#nullable enable

using System.Globalization;
using System.Xml;

namespace ExcelDataReader.Core.XmlSpreadsheetFormat;

internal sealed class SpreadsheetXmlWorksheet : IWorksheet
{
    private const string SpreadsheetNamespace = "urn:schemas-microsoft-com:office:spreadsheet";
    private const string ExcelNamespace = "urn:schemas-microsoft-com:office:excel";

    private readonly List<Row> _rows;
    private readonly Stream? _stream;
    private readonly IReadOnlyDictionary<string, ExtendedFormat>? _stylesById;
    private readonly int _worksheetIndex;

    private SpreadsheetXmlWorksheet(
        string name,
        string visibleState,
        List<Row> rows,
        List<Column> columnWidths,
        CellRange[] mergeCells,
        int fieldCount,
        int rowCount,
        CellRange? dimension,
        Stream? stream = null,
        IReadOnlyDictionary<string, ExtendedFormat>? stylesById = null,
        int worksheetIndex = -1)
    {
        Name = name;
        CodeName = null;
        VisibleState = visibleState;
        HeaderFooter = null;
        _rows = rows;
        ColumnWidths = columnWidths;
        MergeCells = mergeCells;
        FieldCount = fieldCount;
        RowCount = rowCount;
        Dimension = dimension;
        _stream = stream;
        _stylesById = stylesById;
        _worksheetIndex = worksheetIndex;
    }

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
        int expandedColumnCount,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
        bool singlePassMode)
    {
        if (singlePassMode)
        {
            return new SpreadsheetXmlWorksheet(
                name,
                visibleState,
                [],
                [],
                [],
                0,
                0,
                null,
                stream,
                stylesById,
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
                return Parse(workbookReader.ReadSubtree(), stylesById);
            }
        }

        throw new XmlException("Invalid SpreadsheetML worksheet.");
    }

    public static SpreadsheetXmlWorksheet Parse(XmlReader worksheetReader, IReadOnlyDictionary<string, ExtendedFormat> stylesById)
    {
        using (worksheetReader)
        {
            if (!worksheetReader.Read() || worksheetReader.NodeType != XmlNodeType.Element)
                throw new XmlException("Invalid SpreadsheetML worksheet.");

            var name = GetSpreadsheetAttribute(worksheetReader, "Name") ?? string.Empty;
            var visibleState = "visible";
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
                    visibleState = ParseVisibleState(worksheetReader.ReadSubtree());
                }
                else if (worksheetReader.LocalName == "Table" && worksheetReader.NamespaceURI == SpreadsheetNamespace)
                {
                    ParseTable(
                        worksheetReader.ReadSubtree(),
                        stylesById,
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
            foreach (var row in StreamRows(_stream, _worksheetIndex, _stylesById))
                yield return row;
            yield break;
        }

        foreach (var row in _rows)
            yield return row;
    }

    internal static string ParseVisibleState(XmlReader worksheetOptionsReader)
    {
        using (worksheetOptionsReader)
        {
            while (worksheetOptionsReader.Read())
            {
                if (worksheetOptionsReader.NodeType == XmlNodeType.Element &&
                    worksheetOptionsReader.LocalName == "Visible" &&
                    worksheetOptionsReader.NamespaceURI == ExcelNamespace)
                {
                    var visible = worksheetOptionsReader.ReadElementContentAsString();
                    return visible.ToLowerInvariant() switch
                    {
                        "sheethidden" => "hidden",
                        "sheetveryhidden" => "veryhidden",
                        _ => "visible",
                    };
                }
            }
        }

        return "visible";
    }

    private static IEnumerable<Row> StreamRows(
        Stream stream,
        int worksheetIndex,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById)
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

            foreach (var row in StreamRowsFromWorksheet(workbookReader.ReadSubtree(), stylesById))
                yield return row;
            yield break;
        }
    }

    private static IEnumerable<Row> StreamRowsFromWorksheet(
        XmlReader worksheetReader,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById)
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
                    foreach (var row in StreamRowsFromTable(worksheetReader.ReadSubtree(), stylesById))
                        yield return row;
                    yield break;
                }
            }
        }
    }

    private static IEnumerable<Row> StreamRowsFromTable(
        XmlReader tableReader,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById)
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
                    mergeCells,
                    ref currentRowIndex,
                    ref maxColumn,
                    ref maxRow);

                while (row.RowIndex > currentRowIndex)
                {
                    yield return new Row(currentRowIndex, 15D, []);
                    currentRowIndex++;
                }

                yield return row;
                currentRowIndex = row.RowIndex + 1;
            }
        }
    }

    private static Row ParseRowForStreaming(
        XmlReader rowReader,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
        List<CellRange> mergeCells,
        ref int currentRowIndex,
        ref int maxColumn,
        ref int maxRow)
    {
        using (rowReader)
        {
            if (!rowReader.Read() || rowReader.NodeType != XmlNodeType.Element)
                return new Row(currentRowIndex, 15D, []);

            int rowIndexFromAttribute = ParseInt(GetSpreadsheetAttribute(rowReader, "Index"), -1);
            if (rowIndexFromAttribute > 0)
                currentRowIndex = rowIndexFromAttribute - 1;

            bool hidden = ParseBool(GetSpreadsheetAttribute(rowReader, "Hidden"));
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
                    : ExtendedFormat.Zero;

                bool hasData = false;
                var value = ParseCellValue(rowReader.ReadSubtree(), out var cellError, out hasData);

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
        double? width = ParseNullableDouble(GetSpreadsheetAttribute(columnReader, "Width"));
        columnWidths.Add(new Column(currentColumn, currentColumn + span, hidden, width));
        currentColumn += span + 1;
    }

    private static void ParseRow(
        XmlReader rowReader,
        IReadOnlyDictionary<string, ExtendedFormat> stylesById,
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
                    : ExtendedFormat.Zero;

                bool hasData = false;
                var value = ParseCellValue(rowReader.ReadSubtree(), out var cellError, out hasData);

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
            maxRow = Math.Max(maxRow, currentRowIndex);
            currentRowIndex++;
        }
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
                var text = cellReader.ReadElementContentAsString();

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
                if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var dateTime))
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
            IgnoreWhitespace = true,
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
}
