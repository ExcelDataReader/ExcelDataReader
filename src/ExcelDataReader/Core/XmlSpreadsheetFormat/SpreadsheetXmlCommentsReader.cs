using System.Text;
using System.Xml;

namespace ExcelDataReader.Core.XmlSpreadsheetFormat;

internal static class SpreadsheetXmlCommentsReader
{
    private const string Namespace = "urn:schemas-microsoft-com:office:spreadsheet";

    public static Dictionary<long, CellCommentThread> Load(Stream stream, int worksheetIndex)
    {
        var result = new Dictionary<long, CellCommentThread>();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, CloseInput = false });
        int sheet = -1;
        int row = -1;
        int column = -1;
        int nextRow = 0;
        int nextColumn = 0;
        int rowDepth = -1;
        int cellDepth = -1;
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != Namespace)
                continue;
            switch (reader.LocalName)
            {
                case "Worksheet":
                    if (++sheet > worksheetIndex)
                        return result;
                    break;
                case "Row" when sheet == worksheetIndex:
                    row = SpreadsheetXmlWorksheetReader.ReadIndex(reader, nextRow);
                    nextRow = checked(row + 1 + SpreadsheetXmlWorksheetReader.ReadCount(reader, "Span"));
                    nextColumn = 0;
                    rowDepth = reader.Depth;
                    break;
                case "Cell" when sheet == worksheetIndex && reader.Depth == rowDepth + 1:
                    column = SpreadsheetXmlWorksheetReader.ReadIndex(reader, nextColumn);
                    nextColumn = checked(column + 1 + SpreadsheetXmlWorksheetReader.ReadCount(reader, "MergeAcross"));
                    cellDepth = reader.Depth;
                    break;
                case "Comment" when sheet == worksheetIndex && reader.Depth == cellDepth + 1:
                    string? author = reader.GetAttribute("Author", Namespace) ?? reader.GetAttribute("Author");
                    var text = new StringBuilder();
                    using (var comment = reader.ReadSubtree())
                    {
                        while (comment.Read())
                        {
                            if (comment.NodeType != XmlNodeType.Element || comment.LocalName != "Data" || comment.NamespaceURI != Namespace || comment.Depth != 1)
                                continue;
                            using var data = comment.ReadSubtree();
                            while (data.Read())
                            {
                                if (data.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.SignificantWhitespace or XmlNodeType.Whitespace)
                                    text.Append(data.Value);
                            }
                        }
                    }

                    CommentStore.Add(result, row, column, author, text.ToString());
                    break;
            }
        }

        return result;
    }
}