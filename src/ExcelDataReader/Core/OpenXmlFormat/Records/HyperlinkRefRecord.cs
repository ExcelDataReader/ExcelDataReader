namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class HyperlinkRefRecord(string rId, string refAttr) : Record
{
    public string RId { get; } = rId;

    public string RefAttr { get; } = refAttr;
}
