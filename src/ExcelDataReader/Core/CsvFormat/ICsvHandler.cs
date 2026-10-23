namespace ExcelDataReader.Core.CsvFormat;

internal interface ICsvHandler
{
    WorksheetReadControl OnRow(int index);

    WorksheetReadControl OnField(int index, string value);

    WorksheetReadControl OnRowEnd(int fieldCount, bool empty, bool flushing);
}