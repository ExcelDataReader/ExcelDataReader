using System.Globalization;
using System.Text;
using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

internal static class XmlCommentsReader
{
    public static Dictionary<long, CellCommentThread> Load(XmlReader reader)
    {
        const string Transitional = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string Strict = "http://purl.oclc.org/ooxml/spreadsheetml/main";
        var result = new Dictionary<long, CellCommentThread>();
        var authors = new List<string>();
        reader.MoveToContent();
        string ns = reader.NamespaceURI;
        if (reader.LocalName != "comments" || (ns != Transitional && ns != Strict))
            throw new XmlException("Invalid comments part.");
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != ns)
                continue;
            if (reader.LocalName == "author")
            {
                using var author = reader.ReadSubtree();
                author.Read();
                authors.Add(author.ReadElementContentAsString());
            }
            else if (reader.LocalName == "comment")
            {
                string reference = reader.GetAttribute("ref") ?? string.Empty;
                if (reference.Length is < 2 or > 10 ||
                    (reference.Length > 3 && char.IsLetter(reference[3])) ||
                    !ReferenceHelper.ParseReference(reference, out int column, out int row))
                    throw new XmlException("Invalid comment cell reference.");
                int authorId = int.Parse(reader.GetAttribute("authorId") ?? throw new XmlException("Missing comment author."), CultureInfo.InvariantCulture);
                if ((uint)authorId >= (uint)authors.Count)
                    throw new XmlException("Invalid comment author index.");
                var text = new StringBuilder();
                using (var comment = reader.ReadSubtree())
                {
                    while (comment.Read())
                    {
                        if (comment.NodeType == XmlNodeType.Element && comment.NamespaceURI == ns && comment.LocalName == "rPh")
                        {
                            using var phonetic = comment.ReadSubtree();
                            while (phonetic.Read())
                            {
                            }

                            continue;
                        }

                        if (comment.NodeType != XmlNodeType.Element || comment.NamespaceURI != ns || comment.LocalName != "t")
                            continue;
                        using var item = comment.ReadSubtree();
                        item.Read();
                        text.Append(item.ReadElementContentAsString());
                    }
                }

                CommentStore.Add(result, row - 1, column - 1, Helpers.ConvertEscapeChars(authors[authorId]), Helpers.ConvertEscapeChars(text.ToString()));
            }
        }

        return result;
    }
}