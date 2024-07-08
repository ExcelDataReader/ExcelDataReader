namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class CellStyleExtendedFormatRecord(ExtendedFormat extendedFormat) : Record
{
    public ExtendedFormat ExtendedFormat { get; } = extendedFormat;
}
