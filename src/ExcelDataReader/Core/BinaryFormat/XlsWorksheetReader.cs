namespace ExcelDataReader.Core.BinaryFormat;

internal sealed class XlsWorksheetReader : IWorksheetReader
{
    private readonly XlsWorksheet _worksheet;
    private readonly XlsBiffStream _stream;
    private readonly XlsWorksheetCellDecoder _decoder;
    private readonly WorksheetReadMode _mode;
    private readonly List<List<Cell>> _cellListPool = [];
    private readonly Dictionary<int, double> _rowHeights = [];
    private BufferedRow[] _block = [];
    private int _metadataStage;
    private int _metadataIndex;
    private int _blockStart;
    private int _blockCount;
    private int _blockIndex;
    private int _cellIndex;
    private bool _rowOpen;
    private bool _disposed;
    private bool _ended;
    private bool _firstRecord = true;
    private bool _sourceEnded;
    private XlsWorksheetRecord _pendingRecord;
    private bool _hasPendingRecord;
    private int _pendingColumn;
    private ushort? _ixfe;
    private int _currentRow = -1;
    private int _lastCompletedRow = -1;
    private int _maxRow = -1;
    private int _maxColumn = -1;
    private bool _scanResultEmitted;

    public XlsWorksheetReader(XlsWorksheet worksheet, WorksheetReadMode mode = WorksheetReadMode.Materialize)
    {
        _worksheet = worksheet;
        _mode = mode;
        _decoder = new XlsWorksheetCellDecoder(worksheet);
        _stream = new XlsBiffStream(worksheet.Stream, (int)worksheet.DataOffset, worksheet.Workbook.BiffVersion, BIFFTYPE.Worksheet, secretKey: worksheet.Workbook.SecretKey, encryption: worksheet.Workbook.Encryption);
    }

    public WorksheetReadStatus Consume<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        if (_disposed || _ended)
            return WorksheetReadStatus.EndOfInput;

        try
        {
            if (EmitMetadata(ref handler) == WorksheetReadControl.Pause)
                return WorksheetReadStatus.Paused;

            var status = _worksheet.SinglePassMode ? ConsumeForward(ref handler) : ConsumeIndexed(ref handler);
            if (status == WorksheetReadStatus.Paused)
                return status;

            if (_mode == WorksheetReadMode.Scan && !_scanResultEmitted)
            {
                _scanResultEmitted = true;
                if (handler.OnScanResult(_maxRow, _maxColumn) == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
            }

            _ended = true;
            Dispose();
            return WorksheetReadStatus.EndOfInput;
        }
        catch
        {
            Dispose();
            throw;
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        ReleasePendingRecord();
        for (var i = 0; i < _block.Length; i++)
        {
            _block[i].Cells?.Clear();
            _block[i] = default;
        }

        _block = [];
        foreach (var cells in _cellListPool)
            cells.Clear();
        _cellListPool.Clear();
        _rowHeights.Clear();
        _stream.Dispose();
    }

    private WorksheetReadControl EmitMetadata<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        while (_metadataStage < 7)
        {
            WorksheetReadControl control = WorksheetReadControl.Continue;
            switch (_metadataStage)
            {
                case 0:
                    _metadataStage++;
                    if (_worksheet.Dimension is { } dimension)
                        control = handler.OnDimension(dimension);
                    break;
                case 1:
                    _metadataStage++;
                    if (_worksheet.CodeName != null)
                        control = handler.OnCodeName(_worksheet.CodeName);
                    break;
                case 2:
                    _metadataStage++;
                    control = handler.OnDefaultRowHeight(_worksheet.DefaultRowHeight / 20.0);
                    break;
                case 3:
                    if (_metadataIndex < _worksheet.ColumnWidths.Count)
                    {
                        var column = _worksheet.ColumnWidths[_metadataIndex++];
                        control = handler.OnColumn(column);
                    }
                    else
                    {
                        _metadataIndex = 0;
                        _metadataStage++;
                    }

                    break;
                case 4:
                    if (_metadataIndex < _worksheet.MergeCells.Length)
                    {
                        var range = _worksheet.MergeCells[_metadataIndex++];
                        control = handler.OnMergeCell(range);
                    }
                    else
                    {
                        _metadataIndex = 0;
                        _metadataStage++;
                    }

                    break;
                case 5:
                    _metadataStage++;
                    if (_worksheet.HeaderFooter is { } headerFooter)
                        control = handler.OnHeaderFooter(headerFooter);
                    break;
                case 6:
                    _metadataStage++;
                    control = handler.OnSheetDataBegin();
                    break;
            }

            if (control == WorksheetReadControl.Pause)
                return control;
        }

        return WorksheetReadControl.Continue;
    }

    private WorksheetReadStatus ConsumeIndexed<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        while (_blockStart < _worksheet.RowCount)
        {
            if (_blockCount == 0)
                ReadBlock();

            while (_blockIndex < _blockCount)
            {
                ref var row = ref _block[_blockIndex];
                if (!row.Present)
                {
                    _blockIndex++;
                    continue;
                }

                if (!_rowOpen)
                {
                    _rowOpen = true;
                    _maxRow = Math.Max(_maxRow, _blockStart + _blockIndex);
                    if (handler.OnRow(_blockStart + _blockIndex, row.Height == 0, row.Height) == WorksheetReadControl.Pause)
                        return WorksheetReadStatus.Paused;
                }

                var cells = row.Cells;
                if (cells != null)
                {
                    for (var index = _cellIndex; index < cells.Count; index++)
                    {
                        var cell = cells[index];
                        if (handler.OnCell(in cell) == WorksheetReadControl.Pause)
                        {
                            _cellIndex = index + 1;
                            return WorksheetReadStatus.Paused;
                        }
                    }

                    cells.Clear();
                    _cellListPool.Add(cells);
                }

                row = default;
                _cellIndex = 0;
                _blockIndex++;
                _rowOpen = false;
                if (handler.OnRowEnd() == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
            }

            _blockStart += _blockCount;
            _blockCount = 0;
            _blockIndex = 0;
        }

        return WorksheetReadStatus.EndOfInput;
    }

    private void ReadBlock()
    {
        GetBlockSize(out var count, out var minOffset, out var maxOffset);
        _blockCount = count;
        if (_block.Length < count)
            _block = new BufferedRow[count];

        for (var i = 0; i < count; i++)
        {
            if (_worksheet.RowOffsetMap.TryGetValue(_blockStart + i, out var offset))
                InitializeRow(_blockStart + i, offset);
        }

        if (minOffset == int.MaxValue)
            return;

        _stream.Position = minOffset;
        var materialize = _mode == WorksheetReadMode.Materialize;
        ushort? ixfe = null;
        while (_stream.Position <= maxOffset && _stream.ReadWorksheetRecord(out var item))
        {
            try
            {
                if (item.IsCell)
                {
                    var frame = item.Cell;
                    if (materialize)
                    {
                        // The offset map initializes every physical row before this scan.
                        var cells = _block[frame.RowIndex - _blockStart].Cells!;
                        var ownedCell = item.OwnedRecord as XlsBiffBlankCell;
                        if (frame.Id == BIFFRECORDTYPE.MULRK)
                        {
                            var lastColumn = frame.LastColumnIndex;
                            for (var column = (int)frame.ColumnIndex; column <= lastColumn; column++)
                                cells.Add(_decoder.Decode(_stream, frame, ownedCell, ixfe, column));
                        }
                        else
                        {
                            cells.Add(_decoder.Decode(_stream, frame, ownedCell, ixfe, frame.ColumnIndex));
                        }
                    }
                    else
                    {
                        MeasureCell(frame, item.OwnedRecord as XlsBiffBlankCell);
                    }

                    ixfe = null;
                }
                else if (item.OwnedRecord?.Id == BIFFRECORDTYPE.IXFE)
                {
                    ixfe = item.OwnedRecord.ReadUInt16(0);
                }
            }
            finally
            {
                item.Release();
            }
        }
    }

    private void InitializeRow(int index, XlsRowOffset offset)
    {
        ref var row = ref _block[index - _blockStart];
        row.Present = true;
        row.Height = (offset.RowHeight.HasValue && !offset.UseDefaultRowHeight
            ? offset.RowHeight.Value
            : _worksheet.DefaultRowHeight) / 20.0;

        if (_mode == WorksheetReadMode.Materialize)
        {
            if (_cellListPool.Count > 0)
            {
                row.Cells = _cellListPool[^1];
                _cellListPool.RemoveAt(_cellListPool.Count - 1);
            }
            else
            {
                row.Cells = [];
            }
        }
    }

    // Physical overlap may expand this range to the entire sheet. Only this policy
    // buffers normalized cells; the forward policy keeps a single pending frame.
    private void GetBlockSize(out int count, out int minOffset, out int maxOffset)
    {
        count = Math.Min(32, _worksheet.RowCount - _blockStart);
        minOffset = int.MaxValue;
        maxOffset = int.MinValue;
        for (var i = 0; i < count; i++)
        {
            if (!_worksheet.RowOffsetMap.TryGetValue(_blockStart + i, out var offset))
                continue;

            minOffset = Math.Min(minOffset, offset.MinCellOffset);
            maxOffset = Math.Max(maxOffset, offset.MaxCellOffset);
            if (offset.MaxOverlapRowIndex != int.MinValue)
                count = Math.Max(count, offset.MaxOverlapRowIndex + 1 - _blockStart);
        }
    }

    private WorksheetReadStatus ConsumeForward<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        while (true)
        {
            if (!_hasPendingRecord && !_sourceEnded)
                ReadForwardRecord();

            if (_currentRow >= 0 && (_sourceEnded || _pendingRecord.Cell.RowIndex != _currentRow))
            {
                if (!_rowOpen && !_sourceEnded)
                {
                    _rowOpen = true;
                    var height = GetForwardRowHeight();
                    if (handler.OnRow(_currentRow, height == 0, height) == WorksheetReadControl.Pause)
                        return WorksheetReadStatus.Paused;
                }
                else
                {
                    var emitEnd = _rowOpen;
                    _rowHeights.Remove(_currentRow);
                    _lastCompletedRow = _currentRow;
                    _currentRow = -1;
                    _rowOpen = false;
                    if (emitEnd && handler.OnRowEnd() == WorksheetReadControl.Pause)
                        return WorksheetReadStatus.Paused;
                }

                continue;
            }

            if (_sourceEnded)
                return WorksheetReadStatus.EndOfInput;

            var frame = _pendingRecord.Cell;
            if (_currentRow < 0)
            {
                if (frame.RowIndex <= _lastCompletedRow)
                    throw new InvalidOperationException("File contains out-of-order cells. Use SinglePassMode = false to read this file.");

                _currentRow = frame.RowIndex;
            }

            var ownedCell = _pendingRecord.OwnedRecord as XlsBiffBlankCell;
            if (ownedCell is { IsEmpty: true } || (frame.Id == BIFFRECORDTYPE.MULRK && frame.LastColumnIndex < frame.ColumnIndex))
            {
                _ixfe = null;
                ReleasePendingRecord();
                continue;
            }

            if (!_rowOpen)
            {
                _rowOpen = true;
                var height = GetForwardRowHeight();
                if (handler.OnRow(_currentRow, height == 0, height) == WorksheetReadControl.Pause)
                    return WorksheetReadStatus.Paused;
            }

            if (_mode == WorksheetReadMode.Scan)
            {
                _ixfe = null;
                ReleasePendingRecord();
                continue;
            }

            var lastColumn = frame.Id == BIFFRECORDTYPE.MULRK ? frame.LastColumnIndex : frame.ColumnIndex;
            var cell = _decoder.Decode(_stream, frame, ownedCell, _ixfe, _pendingColumn++);
            if (_pendingColumn > lastColumn)
            {
                _ixfe = null;
                ReleasePendingRecord();
            }

            if (handler.OnCell(in cell) == WorksheetReadControl.Pause)
                return WorksheetReadStatus.Paused;
        }
    }

    private double GetForwardRowHeight() =>
        _rowHeights.TryGetValue(_currentRow, out var height) ? height : _worksheet.DefaultRowHeight / 20.0;

    private void ReadForwardRecord()
    {
        while (_stream.ReadWorksheetRecord(out var item))
        {
            if (item.IsCell)
            {
                _firstRecord = false;
                _pendingRecord = item;
                _hasPendingRecord = true;
                _pendingColumn = item.Cell.ColumnIndex;
                MeasureCell(item.Cell, item.OwnedRecord as XlsBiffBlankCell);
                return;
            }

            try
            {
                var record = item.OwnedRecord;
                if (record is XlsBiffEof || (record is XlsBiffBOF && !_firstRecord))
                {
                    _sourceEnded = true;
                    return;
                }

                _firstRecord = false;
                if (record is XlsBiffRow row)
                {
                    _rowHeights[row.RowIndex] = (row.UseDefaultRowHeight ? _worksheet.DefaultRowHeight : row.RowHeight) / 20.0;
                }
                else if (record?.Id == BIFFRECORDTYPE.IXFE)
                {
                    _ixfe = record.ReadUInt16(0);
                }
            }
            finally
            {
                item.Release();
            }
        }

        _sourceEnded = true;
    }

    private void MeasureCell(XlsCellFrame frame, XlsBiffBlankCell? ownedCell)
    {
        _maxRow = Math.Max(_maxRow, frame.RowIndex);
        _maxColumn = Math.Max(_maxColumn, XlsWorksheetCellDecoder.GetLastDefinedColumn(frame, ownedCell));
    }

    private void ReleasePendingRecord()
    {
        if (!_hasPendingRecord)
            return;

        _pendingRecord.Release();
        _pendingRecord = default;
        _hasPendingRecord = false;
    }

    private struct BufferedRow
    {
        public bool Present;
        public double Height;
        public List<Cell>? Cells;
    }
}