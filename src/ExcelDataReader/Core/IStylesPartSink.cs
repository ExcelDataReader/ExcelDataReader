namespace ExcelDataReader.Core;

internal interface IStylesPartSink
{
    void AddExtendedFormat(ExtendedFormat format);

    void AddCellStyleExtendedFormat(ExtendedFormat format);

    void AddNumberFormat(int formatIndexInFile, string formatString);
}