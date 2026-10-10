using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

internal sealed class XmlSharedStringsReader(XmlReader reader, ISharedStringStore store) : IDisposable
{
    private const string ElementSst = "sst";
    private const string ElementStringItem = "si";
    private const string AttributeUniqueCount = "uniqueCount";

    private XmlReader Reader { get; } = reader ?? throw new ArgumentNullException(nameof(reader));

    private ISharedStringStore Store { get; } = store ?? throw new ArgumentNullException(nameof(store));

    private XmlProperNamespaces ProperNamespaces { get; } =
        new(reader.IsStartElement() && reader.NamespaceURI == XmlNamespaces.StrictNsSpreadsheetMl);

    public void Load()
    {
        if (!Reader.IsStartElement(ElementSst, ProperNamespaces.NsSpreadsheetMl))
            return;

        var uniqueCountStr = Reader.GetAttribute(AttributeUniqueCount);
        if (int.TryParse(uniqueCountStr, out var uniqueCount) && uniqueCount > 0)
            Store.Reserve(uniqueCount);

        if (!XmlReaderHelper.ReadFirstContent(Reader))
            return;

        while (!Reader.EOF)
        {
            if (Reader.NodeType == XmlNodeType.Element && Reader.LocalName == ElementStringItem)
            {
                Store.Add(StringHelper.ReadStringItem(Reader, ProperNamespaces.NsSpreadsheetMl));
            }
            else if (!XmlReaderHelper.SkipContent(Reader))
            {
                break;
            }
        }
    }

    public void Dispose() => Reader.Dispose();
}
