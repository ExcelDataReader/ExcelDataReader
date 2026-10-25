using System.Globalization;
using System.Xml;

namespace ExcelDataReader.Core;

internal static class XmlDocumentPropertiesReader
{
    private const string DublinCore = "http://purl.org/dc/elements/1.1/";
    private const string DublinCoreTerms = "http://purl.org/dc/terms/";
    private const string Office = "urn:schemas-microsoft-com:office:office";
    private const string CoreProperties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

    public static void Load(XmlReader reader, DocumentPropertiesBuilder properties, bool spreadsheetXml)
    {
        reader.MoveToContent();
        if (reader.LocalName != (spreadsheetXml ? "DocumentProperties" : "coreProperties") ||
            reader.NamespaceURI != (spreadsheetXml ? Office : CoreProperties))
            return;

        int depth = reader.Depth;
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.Depth != depth + 1)
                continue;

            string name = reader.LocalName;
            string ns = reader.NamespaceURI;
            bool text = spreadsheetXml
                ? ns == Office && name is "Title" or "Subject" or "Author" or "Keywords" or "Description" or "LastAuthor"
                : (ns == DublinCore && name is "title" or "subject" or "creator" or "description") ||
                  (ns == CoreProperties && name is "keywords" or "lastModifiedBy");
            bool date = spreadsheetXml
                ? ns == Office && name is "Created" or "LastSaved" or "LastPrinted"
                : (ns == DublinCoreTerms && name is "created" or "modified") ||
                  (ns == CoreProperties && name == "lastPrinted");
            if (!text && !date)
                continue;

            string? value = ReadSimpleText(reader);
            if (value == null)
                continue;

            switch (name)
            {
                case "Title" or "title": properties.Title = value; break;
                case "Subject" or "subject": properties.Subject = value; break;
                case "Author" or "creator": properties.Author = value; break;
                case "Keywords" or "keywords": properties.Keywords = value; break;
                case "Description" or "description": properties.Description = value; break;
                case "LastAuthor" or "lastModifiedBy": properties.LastModifiedBy = value; break;
                case "Created" or "created": properties.Created = ParseDate(value); break;
                case "LastSaved" or "modified": properties.Modified = ParseDate(value); break;
                case "LastPrinted" or "lastPrinted": properties.LastPrinted = ParseDate(value); break;
            }
        }
    }

    private static DateTime? ParseDate(string value) =>
        DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var date)
            ? date
            : null;

    private static string? ReadSimpleText(XmlReader reader)
    {
        using var subtree = reader.ReadSubtree();
        subtree.Read();
        string value = string.Empty;
        while (subtree.Read())
        {
            if (subtree.NodeType == XmlNodeType.Element)
                return null;
            if (subtree.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.SignificantWhitespace or XmlNodeType.Whitespace)
                value += subtree.Value;
        }

        return value;
    }
}
