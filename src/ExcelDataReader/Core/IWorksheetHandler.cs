namespace ExcelDataReader.Core;

// Callbacks are delivered exactly once, even when they pause consumption.
// Cells own their values; borrowed transport buffers must not escape a callback.
internal interface IWorksheetHandler
{
    WorksheetReadControl OnDimension(CellRange range);

    WorksheetReadControl OnCodeName(string? codeName);

    WorksheetReadControl OnDefaultRowHeight(double? height);

    WorksheetReadControl OnColumn(Column column);

    WorksheetReadControl OnMergeCell(CellRange range);

    WorksheetReadControl OnHeaderFooter(HeaderFooter headerFooter);

    WorksheetReadControl OnSheetDataBegin();

    WorksheetReadControl OnScanResult(int maxRowIndex, int maxColumnIndex);

    WorksheetReadControl OnRow(int index, bool hidden, double? height);

    WorksheetReadControl OnCell(in Cell cell);

    WorksheetReadControl OnRowEnd();
}