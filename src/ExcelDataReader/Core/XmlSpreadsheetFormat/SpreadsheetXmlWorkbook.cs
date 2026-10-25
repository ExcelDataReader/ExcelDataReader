#nullable enable

using System.Globalization;
using System.Xml;

namespace ExcelDataReader.Core.XmlSpreadsheetFormat;

internal sealed class SpreadsheetXmlWorkbook : CommonWorkbook, IWorkbook<SpreadsheetXmlWorksheet>
{
    private const string SpreadsheetNamespace = "urn:schemas-microsoft-com:office:spreadsheet";
    private const string ExcelNamespace = "urn:schemas-microsoft-com:office:excel";

    private readonly WorkbookDescriptorSink _workbookSink = new();
    private readonly Dictionary<string, ExtendedFormat> _stylesById = new(StringComparer.Ordinal)
    {
        ["Default"] = ExtendedFormat.Zero,
    };

    private readonly Stream _stream;

    public SpreadsheetXmlWorkbook(Stream stream)
    {
        _stream = stream;
        ParseWorkbook(stream);
    }

    public int ResultsCount => _workbookSink.Worksheets.Count;

    public int ActiveSheet => _workbookSink.ActiveSheet;

    public DocumentProperties DocumentProperties { get; private set; } = DocumentProperties.Empty;

    public static bool IsSpreadsheetXmlStream(Stream stream)
    {
        if (!stream.CanSeek)
            return false;

        var originalPosition = stream.Position;
        try
        {
            using var reader = CreateXmlReaderAtStart(stream, tolerateLeadingWhitespace: true);
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    return reader.LocalName == "Workbook" && reader.NamespaceURI == SpreadsheetNamespace;
                }
            }

            return false;
        }
        catch (XmlException)
        {
            return false;
        }
        finally
        {
            stream.Seek(originalPosition, SeekOrigin.Begin);
        }
    }

    public IEnumerable<SpreadsheetXmlWorksheet> ReadWorksheets()
    {
        for (int i = 0; i < _workbookSink.Worksheets.Count; i++)
        {
            yield return SpreadsheetXmlWorksheet.Create(
                _stream,
                i,
                _workbookSink.Worksheets[i].Name,
                _workbookSink.Worksheets[i].VisibleState,
                _workbookSink.Worksheets[i].CodeName,
                _workbookSink.Worksheets[i].HeaderFooter,
                _workbookSink.Worksheets[i].ExpandedColumnCount,
                _stylesById,
                Formats,
                SinglePassMode);
        }
    }

    public void Dispose()
    {
        _stream.Dispose();
    }

    private static (string Name, string VisibleState, string? CodeName, HeaderFooter? HeaderFooter, int ExpandedColumnCount) ParseWorksheetDescriptor(XmlReader worksheetReader, XmlReader workbookReader)
    {
        string name = GetSpreadsheetAttribute(workbookReader, "Name") ?? string.Empty;
        string visibleState = "visible";
        string? codeName = null;
        HeaderFooter? headerFooter = null;
        int expandedColumnCount = 0;

        using (worksheetReader)
        {
            while (worksheetReader.Read())
            {
                if (worksheetReader.NodeType != XmlNodeType.Element)
                    continue;

                if (worksheetReader.LocalName == "WorksheetOptions" && worksheetReader.NamespaceURI == ExcelNamespace)
                {
                    (visibleState, codeName, headerFooter) = SpreadsheetXmlWorksheetReader.ParseWorksheetOptions(worksheetReader.ReadSubtree());
                    continue;
                }

                if (worksheetReader.LocalName == "Table" && worksheetReader.NamespaceURI == SpreadsheetNamespace)
                {
                    expandedColumnCount = ParseInt(GetSpreadsheetAttribute(worksheetReader, "ExpandedColumnCount"));
                    SkipElement(worksheetReader);
                }
            }
        }

        return (name, visibleState, codeName, headerFooter, expandedColumnCount);
    }

    private static XmlReader CreateXmlReaderAtStart(Stream stream, bool tolerateLeadingWhitespace)
    {
        if (stream.CanSeek)
        {
            stream.Seek(0, SeekOrigin.Begin);
            if (tolerateLeadingWhitespace)
            {
                SkipLeadingAsciiWhitespace(stream);
            }
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

    private static void SkipElement(XmlReader reader)
    {
        if (reader.IsEmptyElement)
            return;

        int depth = reader.Depth;
        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
        {
        }
    }

    private static string? GetSpreadsheetAttribute(XmlReader reader, string attributeName)
        => reader.GetAttribute(attributeName, SpreadsheetNamespace) ?? reader.GetAttribute(attributeName);

    private static int ParseInt(string? value, int defaultValue = 0)
        => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : defaultValue;

    private static bool ParseBool(string? value)
        => value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);

    private static HorizontalAlignment ParseHorizontalAlignment(string? value) => value switch
    {
        _ when string.Equals(value, "left", StringComparison.OrdinalIgnoreCase) => HorizontalAlignment.Left,
        _ when string.Equals(value, "center", StringComparison.OrdinalIgnoreCase) => HorizontalAlignment.Center,
        _ when string.Equals(value, "right", StringComparison.OrdinalIgnoreCase) => HorizontalAlignment.Right,
        _ when string.Equals(value, "justify", StringComparison.OrdinalIgnoreCase) => HorizontalAlignment.Justified,
        _ when string.Equals(value, "distributed", StringComparison.OrdinalIgnoreCase) => HorizontalAlignment.Distributed,
        _ => HorizontalAlignment.General,
    };

    private static VerticalAlignment ParseVerticalAlignment(string? value) => value switch
    {
        _ when string.Equals(value, "top", StringComparison.OrdinalIgnoreCase) => VerticalAlignment.Top,
        _ when string.Equals(value, "center", StringComparison.OrdinalIgnoreCase) => VerticalAlignment.Center,
        _ when string.Equals(value, "bottom", StringComparison.OrdinalIgnoreCase) => VerticalAlignment.Bottom,
        _ when string.Equals(value, "justify", StringComparison.OrdinalIgnoreCase) => VerticalAlignment.Justify,
        _ when string.Equals(value, "distributed", StringComparison.OrdinalIgnoreCase) => VerticalAlignment.Distributed,
        _ => VerticalAlignment.Bottom,
    };

    private void ParseWorkbook(Stream stream)
    {
        using var reader = CreateXmlReaderAtStart(stream, tolerateLeadingWhitespace: true);

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element)
                continue;

            if (reader.LocalName == "Workbook" && reader.NamespaceURI == SpreadsheetNamespace)
            {
                ParseWorkbookElements(reader);
                return;
            }
        }

        throw new XmlException("Invalid SpreadsheetML workbook.");
    }

    private void ParseWorkbookElements(XmlReader workbookReader)
    {
        while (workbookReader.Read())
        {
            if (workbookReader.NodeType != XmlNodeType.Element)
                continue;

            if (workbookReader.Depth == 1 && workbookReader.LocalName == "DocumentProperties" &&
                workbookReader.NamespaceURI == "urn:schemas-microsoft-com:office:office")
            {
                var properties = new DocumentPropertiesBuilder();
                using var subtree = workbookReader.ReadSubtree();
                XmlDocumentPropertiesReader.Load(subtree, properties, spreadsheetXml: true);
                DocumentProperties = properties.Build();
                continue;
            }

            if (workbookReader.LocalName == "Styles" && workbookReader.NamespaceURI == SpreadsheetNamespace)
            {
                ParseStyles(workbookReader.ReadSubtree(), _stylesById);
                continue;
            }

            if (workbookReader.LocalName == "ExcelWorkbook" && workbookReader.NamespaceURI == ExcelNamespace)
            {
                ParseActiveSheet(workbookReader.ReadSubtree());
                continue;
            }

            if (workbookReader.LocalName == "Worksheet" && workbookReader.NamespaceURI == SpreadsheetNamespace)
            {
                _workbookSink.AddWorksheet(ParseWorksheetDescriptor(workbookReader.ReadSubtree(), workbookReader));
            }
        }
    }

    private void ParseStyles(XmlReader stylesReader, Dictionary<string, ExtendedFormat> stylesById)
    {
        var sink = new StylesSink(this, stylesById);

        using (stylesReader)
        {
            while (stylesReader.Read())
            {
                if (stylesReader.NodeType != XmlNodeType.Element ||
                    stylesReader.LocalName != "Style" ||
                    stylesReader.NamespaceURI != SpreadsheetNamespace)
                {
                    continue;
                }

                sink.AddStyle(stylesReader);
            }
        }
    }

    private void ParseActiveSheet(XmlReader workbookOptionsReader)
    {
        using (workbookOptionsReader)
        {
            while (workbookOptionsReader.Read())
            {
                if (workbookOptionsReader.NodeType == XmlNodeType.Element &&
                    workbookOptionsReader.LocalName == "ActiveSheet" &&
                    workbookOptionsReader.NamespaceURI == ExcelNamespace)
                {
                    var text = workbookOptionsReader.ReadElementContentAsString();
                    if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var activeSheet))
                    {
                        _workbookSink.SetActiveSheet(activeSheet);
                    }

                    return;
                }
            }
        }
    }

    private sealed class WorkbookDescriptorSink
    {
        public List<(string Name, string VisibleState, string? CodeName, HeaderFooter? HeaderFooter, int ExpandedColumnCount)> Worksheets { get; } = [];

        public int ActiveSheet { get; private set; }

        public void AddWorksheet((string Name, string VisibleState, string? CodeName, HeaderFooter? HeaderFooter, int ExpandedColumnCount) worksheet)
            => Worksheets.Add(worksheet);

        public void SetActiveSheet(int activeSheet)
            => ActiveSheet = activeSheet;
    }

    private sealed class StylesSink(SpreadsheetXmlWorkbook workbook, Dictionary<string, ExtendedFormat> stylesById)
    {
        private readonly Dictionary<string, int> _numberFormatIndices = new(StringComparer.Ordinal);
        private int _nextCustomNumberFormat = 164;

        public void AddStyle(XmlReader styleReader)
        {
            var styleId = GetSpreadsheetAttribute(styleReader, "ID") ?? string.Empty;
            if (styleId.Length == 0)
            {
                SkipElement(styleReader);
                return;
            }

            int numberFormatIndex = 0;
            var horizontalAlignment = HorizontalAlignment.General;
            var verticalAlignment = VerticalAlignment.Bottom;
            bool hidden = false;
            bool locked = true;
            ParseStyleContents(
                styleReader.ReadSubtree(),
                ref numberFormatIndex,
                ref horizontalAlignment,
                ref verticalAlignment,
                ref hidden,
                ref locked);

            stylesById[styleId] = new ExtendedFormat(
                parentCellStyleXf: 0,
                fontIndex: 0,
                numberFormatIndex,
                locked,
                hidden,
                indentLevel: 0,
                horizontalAlignment,
                verticalAlignment);
        }

        private void ParseStyleContents(
            XmlReader styleReader,
            ref int numberFormatIndex,
            ref HorizontalAlignment horizontalAlignment,
            ref VerticalAlignment verticalAlignment,
            ref bool hidden,
            ref bool locked)
        {
            using (styleReader)
            {
                while (styleReader.Read())
                {
                    if (styleReader.NodeType != XmlNodeType.Element || styleReader.NamespaceURI != SpreadsheetNamespace)
                        continue;

                    if (styleReader.LocalName == "NumberFormat")
                    {
                        var numberFormat = GetSpreadsheetAttribute(styleReader, "Format") ?? string.Empty;
                        if (numberFormat.Length == 0)
                            continue;

                        if (_numberFormatIndices.TryGetValue(numberFormat, out var existingIndex))
                        {
                            numberFormatIndex = existingIndex;
                        }
                        else
                        {
                            numberFormatIndex = _nextCustomNumberFormat++;
                            _numberFormatIndices[numberFormat] = numberFormatIndex;
                            workbook.AddNumberFormat(numberFormatIndex, numberFormat);
                        }
                    }
                    else if (styleReader.LocalName == "Alignment")
                    {
                        horizontalAlignment = ParseHorizontalAlignment(GetSpreadsheetAttribute(styleReader, "Horizontal"));
                        verticalAlignment = ParseVerticalAlignment(GetSpreadsheetAttribute(styleReader, "Vertical"));
                    }
                    else if (styleReader.LocalName == "Protection")
                    {
                        hidden = ParseBool(GetSpreadsheetAttribute(styleReader, "Hidden"));
                        locked = !string.Equals(GetSpreadsheetAttribute(styleReader, "Protected"), "0", StringComparison.OrdinalIgnoreCase);
                    }
                }
            }
        }
    }
}