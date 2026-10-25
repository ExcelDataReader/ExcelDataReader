using System.Xml;

namespace ExcelDataReader.Core.XmlSpreadsheetFormat;

internal static class SpreadsheetXmlHyperlinksReader
{
    private const string SpreadsheetNamespace = "urn:schemas-microsoft-com:office:spreadsheet";
    private const string ExcelNamespace = "urn:schemas-microsoft-com:office:excel";

    public static HyperlinkStore Load(Stream stream, int worksheetIndex)
    {
        var result = new HyperlinkStore();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, CloseInput = false });
        int sheet = -1;
        int sheetDepth = -1;
        int tableDepth = -1;
        int rowDepth = -1;
        int row = -1;
        int nextRow = 0;
        int nextColumn = 0;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement)
            {
                if (reader.Depth == rowDepth)
                    rowDepth = -1;
                if (reader.Depth == tableDepth)
                    tableDepth = -1;
                if (reader.Depth == sheetDepth && sheet == worksheetIndex)
                    return result;
                continue;
            }

            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != SpreadsheetNamespace)
                continue;
            switch (reader.LocalName)
            {
                case "Worksheet" when reader.Depth == 1:
                    sheet++;
                    sheetDepth = reader.Depth;
                    if (sheet > worksheetIndex)
                        return result;
                    break;
                case "Table" when sheet == worksheetIndex && reader.Depth == sheetDepth + 1 && !reader.IsEmptyElement:
                    tableDepth = reader.Depth;
                    break;
                case "Row" when sheet == worksheetIndex && tableDepth >= 0 && reader.Depth == tableDepth + 1:
                    row = SpreadsheetXmlWorksheetReader.ReadIndex(reader, nextRow);
                    nextRow = checked(row + 1 + SpreadsheetXmlWorksheetReader.ReadCount(reader, "Span"));
                    nextColumn = 0;
                    rowDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                    break;
                case "Cell" when sheet == worksheetIndex && rowDepth >= 0 && reader.Depth == rowDepth + 1:
                    int column = SpreadsheetXmlWorksheetReader.ReadIndex(reader, nextColumn);
                    nextColumn = checked(column + 1 + SpreadsheetXmlWorksheetReader.ReadCount(reader, "MergeAcross"));
                    string? address = reader.GetAttribute("HRef", SpreadsheetNamespace) ?? reader.GetAttribute("HRef");
                    string? tooltip = reader.GetAttribute("HRefScreenTip", ExcelNamespace);
                    if (address != null || tooltip != null)
                        result.Add(new CellRange(column, row, column, row), new CellHyperlink(address ?? string.Empty, tooltip));
                    break;
            }
        }

        return result;
    }
}
