namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class ExtendedFormatRecord(ExtendedFormat extendedFormat) : Record
{
    public ExtendedFormat ExtendedFormat { get; } = extendedFormat;
}
