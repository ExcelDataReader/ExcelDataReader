#nullable enable

using System.Globalization;
using System.Xml;

namespace ExcelDataReader.Core.XmlSpreadsheetFormat;

internal sealed class SpreadsheetXmlWorkbook : CommonWorkbook, IWorkbook<SpreadsheetXmlWorksheet>
{
    private const string SpreadsheetNamespace = "urn:schemas-microsoft-com:office:spreadsheet";
    private const string ExcelNamespace = "urn:schemas-microsoft-com:office:excel";

    private readonly List<(string Name, string VisibleState, string? CodeName, HeaderFooter? HeaderFooter, int ExpandedColumnCount)> _worksheets = [];
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

    public int ResultsCount => _worksheets.Count;

    public int ActiveSheet { get; private set; }

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
        catch
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
        for (int i = 0; i < _worksheets.Count; i++)
        {
            yield return SpreadsheetXmlWorksheet.Create(
                _stream,
                i,
                _worksheets[i].Name,
                _worksheets[i].VisibleState,
                _worksheets[i].CodeName,
                _worksheets[i].HeaderFooter,
                _worksheets[i].ExpandedColumnCount,
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
                    (visibleState, codeName, headerFooter) = SpreadsheetXmlWorksheet.ParseWorksheetOptions(worksheetReader.ReadSubtree());
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

    private static HorizontalAlignment ParseHorizontalAlignment(string? value) => value?.ToLowerInvariant() switch
    {
        "left" => HorizontalAlignment.Left,
        "center" => HorizontalAlignment.Center,
        "right" => HorizontalAlignment.Right,
        "justify" => HorizontalAlignment.Justified,
        "distributed" => HorizontalAlignment.Distributed,
        _ => HorizontalAlignment.General,
    };

    private static VerticalAlignment ParseVerticalAlignment(string? value) => value?.ToLowerInvariant() switch
    {
        "top" => VerticalAlignment.Top,
        "center" => VerticalAlignment.Center,
        "bottom" => VerticalAlignment.Bottom,
        "justify" => VerticalAlignment.Justify,
        "distributed" => VerticalAlignment.Distributed,
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
                _worksheets.Add(ParseWorksheetDescriptor(workbookReader.ReadSubtree(), workbookReader));
            }
        }
    }

    private void ParseStyles(XmlReader stylesReader, Dictionary<string, ExtendedFormat> stylesById)
    {
        var numberFormatIndices = new Dictionary<string, int>(StringComparer.Ordinal);
        int nextCustomNumberFormat = 164;

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

                var styleId = GetSpreadsheetAttribute(stylesReader, "ID") ?? string.Empty;
                if (styleId.Length == 0)
                {
                    SkipElement(stylesReader);
                    continue;
                }

                int numberFormatIndex = 0;
                var horizontalAlignment = HorizontalAlignment.General;
                var verticalAlignment = VerticalAlignment.Bottom;
                bool hidden = false;
                bool locked = true;

                ParseStyleContents(
                    stylesReader.ReadSubtree(),
                    ref numberFormatIndex,
                    ref horizontalAlignment,
                    ref verticalAlignment,
                    ref hidden,
                    ref locked,
                    ref nextCustomNumberFormat,
                    numberFormatIndices);

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
        }
    }

    private void ParseStyleContents(
        XmlReader styleReader,
        ref int numberFormatIndex,
        ref HorizontalAlignment horizontalAlignment,
        ref VerticalAlignment verticalAlignment,
        ref bool hidden,
        ref bool locked,
        ref int nextCustomNumberFormat,
        Dictionary<string, int> numberFormatIndices)
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
                    if (numberFormat.Length > 0)
                    {
                        if (numberFormatIndices.TryGetValue(numberFormat, out var existingIndex))
                        {
                            numberFormatIndex = existingIndex;
                        }
                        else
                        {
                            numberFormatIndex = nextCustomNumberFormat++;
                            numberFormatIndices[numberFormat] = numberFormatIndex;
                            AddNumberFormat(numberFormatIndex, numberFormat);
                        }
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
                        ActiveSheet = activeSheet;
                    }

                    return;
                }
            }
        }
    }
}
