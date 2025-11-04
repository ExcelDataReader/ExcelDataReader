using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

/// <summary>
/// Reads relationship records from a .rels XML file.
/// </summary>
internal sealed class XmlRelReader(XmlReader reader) : XmlRecordReader(reader)
{
    private const string ElementRelationship = "Relationship";
    private const string AttributeId = "Id";
    private const string AttributeType = "Type";
    private const string AttributeTarget = "Target";

    protected override IEnumerable<Record> ReadOverride()
    {
        while (Reader.Read())
        {
            if (Reader.NodeType == XmlNodeType.Element && Reader.LocalName == ElementRelationship)
            {
                var id = Reader.GetAttribute(AttributeId);
                var type = Reader.GetAttribute(AttributeType);
                var target = Reader.GetAttribute(AttributeTarget);

                yield return new RelRecord(id, type, target);
            }
        }
    }
}
