namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class ExtendedFormatRecord : Record
{
    public ExtendedFormatRecord(ExtendedFormat extendedFormat) 
    {
        ExtendedFormat = extendedFormat;
    }

    public ExtendedFormat ExtendedFormat { get; }
}
