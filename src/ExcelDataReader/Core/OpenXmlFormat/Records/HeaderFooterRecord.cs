namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class HeaderFooterRecord(HeaderFooter headerFooter) : Record
{
    public HeaderFooter HeaderFooter { get; } = headerFooter;
}
