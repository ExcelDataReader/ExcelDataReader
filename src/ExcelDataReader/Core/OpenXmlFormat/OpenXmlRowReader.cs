namespace ExcelDataReader.Core.OpenXmlFormat;

internal sealed class OpenXmlRowReader(IWorksheetReader reader, double defaultRowHeight, OpenXmlRowOrder.Region[] regions) : IWorksheetReader
{
    private readonly double _defaultRowHeight = defaultRowHeight;
    private int _rowIndex;
    private int _physicalRowIndex;
    private bool _hidden;
    private double? _height;
    private bool _pendingHeader;
    private bool _rowOpen;
    private bool _endOfInput;
    private int _physicalOrdinal = -1;
    private int _regionIndex;
    private bool _buffering;
    private bool _flushing;
    private List<BufferedRow>? _rows;
    private List<BufferedCell>? _cells;
    private Dictionary<int, int>? _rowLookup;
    private int _bufferRow;
    private int _flushRow;
    private int _flushCell;
    private FlushStage _flushStage;

    private enum FlushStage
    {
        LeadingGapBegin,
        LeadingGapEnd,
        RowBegin,
        Cells,
    }

    public WorksheetReadStatus Consume<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        while (true)
        {
            if (_buffering && (_endOfInput || (_pendingHeader && _physicalOrdinal > regions[_regionIndex].End)))
            {
                _buffering = false;
                _flushing = true;
                _regionIndex++;
                _rows!.Sort(static (left, right) => left.Index.CompareTo(right.Index));
                _flushRow = 0;
                _flushStage = _rowIndex == 0 && _rows[0].Index > 0 ? FlushStage.LeadingGapBegin : FlushStage.RowBegin;
            }

            if (_flushing && Flush(ref handler) == WorksheetReadControl.Pause)
                return WorksheetReadStatus.Paused;

            if (_pendingHeader)
            {
                if (!_buffering && _rowOpen && _rowIndex != _physicalRowIndex)
                {
                    _rowOpen = false;
                    _rowIndex++;
                    if (handler.OnRowEnd() == WorksheetReadControl.Pause)
                        return WorksheetReadStatus.Paused;
                }

                if (!_buffering && _regionIndex < regions.Length && _physicalOrdinal == regions[_regionIndex].Start)
                {
                    _buffering = true;
                    _rows ??= [];
                    _cells ??= [];
                    _rowLookup ??= [];
                }

                if (_buffering)
                {
                    BufferHeader();
                    _pendingHeader = false;
                }
                else
                {
                    if (_physicalRowIndex < _rowIndex)
                        throw new InvalidOperationException("Worksheet rows are out of order. Use SinglePassMode = false to read non-sequential rows.");

                    if (!_rowOpen && _rowIndex == 0 && _physicalRowIndex != 0)
                    {
                        _rowOpen = true;
                        if (handler.OnRow(0, false, 0) == WorksheetReadControl.Pause)
                            return WorksheetReadStatus.Paused;
                        continue;
                    }

                    _rowIndex = _physicalRowIndex;
                    _pendingHeader = false;
                    _rowOpen = true;
                    if (handler.OnRow(_rowIndex, _hidden, _height ?? _defaultRowHeight) == WorksheetReadControl.Pause)
                        return WorksheetReadStatus.Paused;
                }
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

    public void Dispose()
    {
        _rows?.Clear();
        _cells?.Clear();
        _rowLookup?.Clear();
        reader.Dispose();
    }

    private void BufferHeader()
    {
        if (!_rowLookup!.TryGetValue(_physicalRowIndex, out _bufferRow))
        {
            _bufferRow = _rows!.Count;
            _rowLookup.Add(_physicalRowIndex, _bufferRow);
            _rows.Add(new BufferedRow(_physicalRowIndex, _hidden, _height, -1, -1));
        }
        else
        {
            var row = _rows![_bufferRow];
            _rows[_bufferRow] = row with { Hidden = _hidden, Height = _height };
        }
    }

    private void BufferCell(in Cell cell)
    {
        var row = _rows![_bufferRow];
        int index = _cells!.Count;
        _cells.Add(new BufferedCell(cell, -1));
        if (row.Last >= 0)
            _cells[row.Last] = _cells[row.Last] with { Next = index };
        _rows[_bufferRow] = row with { First = row.First < 0 ? index : row.First, Last = index };
    }

    private WorksheetReadControl Flush<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        while (_flushRow < _rows!.Count)
        {
            if (_flushStage == FlushStage.LeadingGapBegin)
            {
                _flushStage = FlushStage.LeadingGapEnd;
                if (handler.OnRow(0, false, 0) == WorksheetReadControl.Pause)
                    return WorksheetReadControl.Pause;
            }

            if (_flushStage == FlushStage.LeadingGapEnd)
            {
                _flushStage = FlushStage.RowBegin;
                if (handler.OnRowEnd() == WorksheetReadControl.Pause)
                    return WorksheetReadControl.Pause;
            }

            var row = _rows[_flushRow];
            if (_flushStage == FlushStage.RowBegin)
            {
                _flushStage = FlushStage.Cells;
                _flushCell = row.First;
                if (handler.OnRow(row.Index, row.Hidden, row.Height ?? _defaultRowHeight) == WorksheetReadControl.Pause)
                    return WorksheetReadControl.Pause;
            }

            while (_flushCell >= 0)
            {
                var cell = _cells![_flushCell];
                _cells[_flushCell] = default;
                _flushCell = cell.Next;
                var value = cell.Cell;
                if (handler.OnCell(in value) == WorksheetReadControl.Pause)
                    return WorksheetReadControl.Pause;
            }

            _flushRow++;
            _flushStage = FlushStage.RowBegin;
            _rowIndex = row.Index + 1;
            if (handler.OnRowEnd() == WorksheetReadControl.Pause)
                return WorksheetReadControl.Pause;
        }

        _rows.Clear();
        _cells!.Clear();
        _rowLookup!.Clear();
        _flushing = false;
        return WorksheetReadControl.Continue;
    }

    private readonly record struct BufferedRow(int Index, bool Hidden, double? Height, int First, int Last);

    private readonly record struct BufferedCell(Cell Cell, int Next);

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
            owner._physicalOrdinal++;
            owner._physicalRowIndex = index;
            owner._hidden = hidden;
            owner._height = height;
            if (!owner._buffering && owner._rowOpen && owner._rowIndex == index)
            {
                return Handler.OnRow(index, hidden, height ?? owner._defaultRowHeight);
            }

            owner._pendingHeader = true;
            return WorksheetReadControl.Pause;
        }

        public WorksheetReadControl OnCell(in Cell cell)
        {
            if (!owner._buffering)
                return Handler.OnCell(in cell);

            owner.BufferCell(in cell);
            return WorksheetReadControl.Continue;
        }

        public readonly WorksheetReadControl OnRowEnd() => WorksheetReadControl.Continue;
    }
}