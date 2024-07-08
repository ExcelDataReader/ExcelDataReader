namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class SheetFormatPrRecord(double? defaultRowHeight) : Record
{
    public double? DefaultRowHeight { get; } = defaultRowHeight;
}
