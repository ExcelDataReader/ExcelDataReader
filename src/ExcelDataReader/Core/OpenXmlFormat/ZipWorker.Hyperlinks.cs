using System.IO.Compression;
using System.Xml;
using ExcelDataReader.Core;
using ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;
using ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

namespace ExcelDataReader.Core.OpenXmlFormat;

internal sealed partial class ZipWorker
{
    private const string TransitionalHyperlinkRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    private const string StrictHyperlinkRelationship = "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink";
    private const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    public HyperlinkStore LoadHyperlinks(string worksheetPath)
    {
        worksheetPath = worksheetPath.TrimStart('/');
        var worksheetEntry = FindEntry(worksheetPath);
        if (worksheetEntry == null)
            return new HyperlinkStore();

        string[] pathParts = worksheetPath.Split('/');
        string directory = pathParts.Length <= 1 ? string.Empty : string.Join("/", pathParts, 0, pathParts.Length - 1) + "/";
        string relationshipPath = directory + "_rels/" + pathParts[^1] + ".rels";
        var relationships = FindEntry(relationshipPath) is { } relationshipEntry
            ? LoadHyperlinkRelationships(relationshipEntry)
            : new Dictionary<string, string>(StringComparer.Ordinal);

        if (worksheetEntry.FullName.EndsWith(".bin", StringComparison.OrdinalIgnoreCase))
        {
            using var reader = new BiffHyperlinksReader(OpenZipEntry(worksheetEntry), relationships);
            return reader.Load();
        }

        if (worksheetEntry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
        {
            using var stream = OpenZipEntry(worksheetEntry);
            using var xml = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, IgnoreWhitespace = true });
            return XmlHyperlinksReader.Load(xml, relationships);
        }

        return new HyperlinkStore();
    }

    private static Dictionary<string, string> LoadHyperlinkRelationships(ZipArchiveEntry entry)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, IgnoreWhitespace = true });
        reader.MoveToContent();
        if (reader.LocalName != "Relationships" || reader.NamespaceURI != PackageRelationships)
            throw new XmlException("Invalid worksheet relationships part.");

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.Depth != 1 ||
                reader.LocalName != "Relationship" || reader.NamespaceURI != PackageRelationships)
                continue;

            string? id = reader.GetAttribute("Id");
            if (string.IsNullOrEmpty(id))
                continue;

            // Duplicate IDs are invalid OPC, but must not reuse an earlier target
            // when the later relationship is missing its target or has another type.
            result.Remove(id);
            string? type = reader.GetAttribute("Type");
            if (type is not (TransitionalHyperlinkRelationship or StrictHyperlinkRelationship))
                continue;

            string? target = reader.GetAttribute("Target");
            if (target != null)
                result[id] = target;
        }

        return result;
    }
}
