using System.Globalization;
using System.Xml;

namespace ExcelDataReader.Core;

internal sealed class DocumentPropertiesBuilder
{
    private const string DublinCore = "http://purl.org/dc/elements/1.1/";
    private const string DublinCoreTerms = "http://purl.org/dc/terms/";
    private const string Office = "urn:schemas-microsoft-com:office:office";
    private const string CoreProperties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

    public string? Title { get; set; }

    public string? Subject { get; set; }

    public string? Author { get; set; }

    public string? Keywords { get; set; }

    public string? Description { get; set; }

    public string? LastModifiedBy { get; set; }

    public DateTime? Created { get; set; }

    public DateTime? Modified { get; set; }

    public DateTime? LastPrinted { get; set; }

    public DocumentProperties Build() => new()
    {
        Title = Title,
        Subject = Subject,
        Author = Author,
        Keywords = Keywords,
        Description = Description,
        LastModifiedBy = LastModifiedBy,
        Created = Created,
        Modified = Modified,
        LastPrinted = LastPrinted,
    };

    public void ReadXml(XmlReader reader, bool spreadsheetXml)
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
                case "Title" or "title": Title = value; break;
                case "Subject" or "subject": Subject = value; break;
                case "Author" or "creator": Author = value; break;
                case "Keywords" or "keywords": Keywords = value; break;
                case "Description" or "description": Description = value; break;
                case "LastAuthor" or "lastModifiedBy": LastModifiedBy = value; break;
                case "Created" or "created": Created = ParseDate(value); break;
                case "LastSaved" or "modified": Modified = ParseDate(value); break;
                case "LastPrinted" or "lastPrinted": LastPrinted = ParseDate(value); break;
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
