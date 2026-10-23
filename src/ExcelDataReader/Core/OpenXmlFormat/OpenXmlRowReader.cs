namespace ExcelDataReader.Core.OpenXmlFormat;

// Consecutive physical row fragments share the current output buffer. A different
// row header completes it; previously returned rows are never reopened.
internal sealed class OpenXmlRowReader(IWorksheetReader reader, double defaultRowHeight) : IWorksheetReader
{
    private int _rowIndex;
    private int _physicalRowIndex;
    private bool _hidden;
    private double? _height;
    private bool _pendingHeader;
    private bool _rowOpen;
    private bool _endOfInput;

    public WorksheetReadStatus Consume<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        while (true)
        {
            if (_pendingHeader)
            {
                if (_rowOpen && _rowIndex != _physicalRowIndex)
                {
                    _rowOpen = false;
                    _rowIndex++;
                    if (handler.OnRowEnd() == WorksheetReadControl.Pause)
                        return WorksheetReadStatus.Paused;
                }

                if (!_rowOpen && _rowIndex == 0 && _physicalRowIndex != 0)
                {
                    _rowOpen = true;
                    if (handler.OnRow(0, false, 0) == WorksheetReadControl.Pause)
                        return WorksheetReadStatus.Paused;
                    continue;
                }

                _rowIndex = Math.Max(_rowIndex, _physicalRowIndex);
                _pendingHeader = false;
                _rowOpen = true;
                if (handler.OnRow(_rowIndex, _hidden, _height ?? defaultRowHeight) == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
            }

            if (_endOfInput)
            {
                if (_rowOpen)
                {
                    _rowOpen = false;
                    if (handler.OnRowEnd() == WorksheetReadControl.Pause)
                        return WorksheetReadStatus.Paused;
                }

                return WorksheetReadStatus.EndOfInput;
            }

            var forwarder = new FragmentHandler<THandler>(this, handler);
            WorksheetReadStatus status;
            try
            {
                status = reader.Consume(ref forwarder);
            }
            finally
            {
                handler = forwarder.Handler;
            }

            if (status == WorksheetReadStatus.EndOfInput)
                _endOfInput = true;
            else if (!_pendingHeader)
                return WorksheetReadStatus.Paused;
        }
    }

    public void Dispose() => reader.Dispose();

    private struct FragmentHandler<THandler>(OpenXmlRowReader owner, THandler handler) : IWorksheetHandler
        where THandler : struct, IWorksheetHandler
    {
        public THandler Handler = handler;

        public WorksheetReadControl OnDimension(CellRange range) => Handler.OnDimension(range);

        public WorksheetReadControl OnCodeName(string? codeName) => Handler.OnCodeName(codeName);

        public WorksheetReadControl OnDefaultRowHeight(double? height) => Handler.OnDefaultRowHeight(height);

        public WorksheetReadControl OnColumn(Column column) => Handler.OnColumn(column);

        public WorksheetReadControl OnMergeCell(CellRange range) => Handler.OnMergeCell(range);

        public WorksheetReadControl OnHeaderFooter(HeaderFooter headerFooter) => Handler.OnHeaderFooter(headerFooter);

        public WorksheetReadControl OnSheetDataBegin() => Handler.OnSheetDataBegin();

        public WorksheetReadControl OnScanResult(int maxRowIndex, int maxColumnIndex) => Handler.OnScanResult(maxRowIndex, maxColumnIndex);

        public WorksheetReadControl OnRow(int index, bool hidden, double? height)
        {
            owner._physicalRowIndex = index;
            owner._hidden = hidden;
            owner._height = height;
            owner._pendingHeader = true;
            return WorksheetReadControl.Pause;
        }

        public WorksheetReadControl OnCell(in Cell cell) => Handler.OnCell(in cell);

        public WorksheetReadControl OnRowEnd() => WorksheetReadControl.Continue;
    }
}
