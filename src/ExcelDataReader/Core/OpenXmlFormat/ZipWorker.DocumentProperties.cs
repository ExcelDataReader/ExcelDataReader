using System.Xml;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.OpenXmlFormat;

internal sealed partial class ZipWorker
{
    public DocumentProperties LoadDocumentProperties()
    {
        var properties = new DocumentPropertiesBuilder();
        try
        {
            string path = "docProps/core.xml";
            if (FindEntry("_rels/.rels") is { } relationships)
            {
                using var stream = relationships.Open();
                using var reader = XmlReader.Create(stream, XmlSettings);
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element || reader.Depth != 1 ||
                        reader.LocalName != "Relationship" ||
                        reader.NamespaceURI != "http://schemas.openxmlformats.org/package/2006/relationships" ||
                        reader.GetAttribute("Type") != "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
                        continue;

                    if (reader.GetAttribute("TargetMode") == "External" || reader.GetAttribute("Target") is not { Length: > 0 } target)
                        return DocumentProperties.Empty;

                    try
                    {
                        path = ResolveCommentPath(string.Empty, target);
                    }
                    catch (ExcelReaderException)
                    {
                        // Invalid property-part paths are optional metadata, not worksheet corruption.
                        return DocumentProperties.Empty;
                    }

                    break;
                }
            }

            if (FindEntry(path) is { } entry)
            {
                using var stream = entry.Open();
                using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
                properties.ReadXml(reader, spreadsheetXml: false);
            }
        }
        catch (XmlException)
        {
            // Optional document properties must not prevent reading the workbook.
        }

        return properties.Build();
    }
}
