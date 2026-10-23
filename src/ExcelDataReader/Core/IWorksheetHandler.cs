namespace ExcelDataReader.Core;

// Callbacks are delivered exactly once, even when they pause consumption.
// A pause accepts the current callback: readers advance their state before calling
// the handler and resume with the next event, not by replaying the paused event.
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

    // Raw parsers report physical rows. A logical-row scheduler may repeat OnRow
    // for fragments of its current row: update metadata without clearing its cells.
    // Scan mode suppresses values; structural callbacks vary by format.
    WorksheetReadControl OnRow(int index, bool hidden, double? height);

    WorksheetReadControl OnCell(in Cell cell);

    // Raw row ends delimit physical fragments; the scheduler replaces them with
    // one completion callback for each assembled logical row.
    WorksheetReadControl OnRowEnd();
}