using System.Xml;
using ExcelDataReader.Core;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

internal static class XmlHyperlinksReader
{
    public static HyperlinkStore Load(XmlReader reader, IReadOnlyDictionary<string, string> relationships)
    {
        var result = new HyperlinkStore();
        reader.MoveToContent();
        string ns = reader.NamespaceURI;
        if (reader.LocalName != "worksheet" || (ns != XmlNamespaces.NsSpreadsheetMl && ns != XmlNamespaces.StrictNsSpreadsheetMl))
            throw new XmlException("Invalid worksheet part.");

        bool inHyperlinks = false;
        int hyperlinksDepth = -1;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && inHyperlinks && reader.Depth == hyperlinksDepth)
            {
                inHyperlinks = false;
                continue;
            }

            if (reader.NodeType != XmlNodeType.Element)
                continue;

            if (reader is { Depth: 1, LocalName: "hyperlinks" } && reader.NamespaceURI == ns)
            {
                inHyperlinks = !reader.IsEmptyElement;
                hyperlinksDepth = reader.Depth;
                continue;
            }

            if (!inHyperlinks || reader.Depth != hyperlinksDepth + 1 || reader.LocalName != "hyperlink" || reader.NamespaceURI != ns)
                continue;

            string reference = reader.GetAttribute("ref") ?? string.Empty;
            CellRange range = HyperlinkStore.ParseRange(reference);
            string? relationshipId = reader.GetAttribute("id", ns == XmlNamespaces.StrictNsSpreadsheetMl ? XmlNamespaces.StrictNsDocumentRelationship : XmlNamespaces.NsDocumentRelationship);
            string? location = reader.GetAttribute("location");
            string? tooltip = reader.GetAttribute("tooltip");

            string? address;
            if (relationshipId != null)
            {
                if (relationshipId.Length == 0 || !relationships.TryGetValue(relationshipId, out var target))
                    continue;
                address = HyperlinkStore.JoinAddress(target, location);
            }
            else
            {
                address = HyperlinkStore.JoinAddress(null, location);
                if (string.IsNullOrEmpty(location) && tooltip == null)
                    continue;
            }

            result.Add(range, new CellHyperlink(address, string.IsNullOrEmpty(tooltip) ? null : tooltip));
        }

        return result;
    }
}
