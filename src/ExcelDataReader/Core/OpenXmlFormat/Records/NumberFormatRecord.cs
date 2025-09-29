namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class NumberFormatRecord(int formatIndexInFile, string formatString) : Record
{
    public int FormatIndexInFile { get; } = formatIndexInFile;

    public string FormatString { get; } = formatString;
}
