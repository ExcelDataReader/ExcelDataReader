using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;
using ExcelDataReader.Core.OpenXmlFormat.XmlFormat;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.OpenXmlFormat;

internal sealed partial class ZipWorker
{
    public Dictionary<long, CellCommentThread>? LoadComments(string worksheetPath)
    {
        worksheetPath = worksheetPath.TrimStart('/');
        int slash = worksheetPath.LastIndexOf('/');
        string directory = slash < 0 ? string.Empty : worksheetPath[..(slash + 1)];
        string name = worksheetPath[(slash + 1)..];
        var relationships = FindEntry(directory + "_rels/" + name + ".rels");
        if (relationships == null)
            return null;

        using var stream = relationships.Open();
        using var reader = XmlReader.Create(stream, XmlSettings);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship" ||
                reader.NamespaceURI != "http://schemas.openxmlformats.org/package/2006/relationships")
                continue;
            string? type = reader.GetAttribute("Type");
            if (type is not ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" or
                "http://purl.oclc.org/ooxml/officeDocument/relationships/comments"))
                continue;
            if (reader.GetAttribute("TargetMode") == "External")
                continue;
            string target = reader.GetAttribute("Target") ?? throw new ExcelReaderException("Missing comments relationship target.");
            var entry = FindEntry(ResolveCommentPath(directory, target)) ?? throw new ExcelReaderException("Missing comments part.");
            using var comments = OpenZipEntry(entry);
            if (entry.FullName.EndsWith(".bin", StringComparison.OrdinalIgnoreCase))
            {
                using var binary = new BiffCommentsReader(comments);
                return binary.Load();
            }

            using var xml = XmlReader.Create(comments, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, IgnoreWhitespace = false });
            return XmlCommentsReader.Load(xml);
        }

        return null;
    }

    private static string ResolveCommentPath(string directory, string target)
    {
        var parts = new List<string>();
        foreach (string part in (target.Length > 0 && target[0] == '/' ? target : directory + target).Split('/'))
        {
            if (part is "" or ".")
                continue;
            if (part == "..")
            {
                if (parts.Count == 0)
                    throw new ExcelReaderException("Invalid comments relationship target.");
                parts.RemoveAt(parts.Count - 1);
            }
            else
            {
                parts.Add(Uri.UnescapeDataString(part));
            }
        }

        return string.Join("/", parts);
    }
}