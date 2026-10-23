namespace ExcelDataReader.Core.BinaryFormat;

internal sealed class XlsWorksheetReader : IWorksheetReader
{
    private readonly XlsWorksheet _worksheet;
    private readonly XlsBiffStream _stream;
    private readonly XlsWorksheetCellDecoder _decoder;
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
    private XlsBiffFrame _pendingFrame;
    private bool _hasPendingRecord;
    private int _pendingColumn;
    private ushort? _ixfe;
    private int _currentRow = -1;
    private int _lastCompletedRow = -1;

    public XlsWorksheetReader(XlsWorksheet worksheet)
    {
        _worksheet = worksheet;
        _decoder = new XlsWorksheetCellDecoder(worksheet);
        _stream = XlsBiffStartup.Open(worksheet.Stream, out _, out _, (int)worksheet.DataOffset, worksheet.Workbook.BiffVersion, BIFFTYPE.Worksheet, secretKey: worksheet.Workbook.SecretKey, encryption: worksheet.Workbook.Encryption);
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
        ushort? ixfe = null;
        while (_stream.Position <= maxOffset && _stream.ReadFrame(out var frame))
        {
            if (frame.IsCell)
            {
                // The offset map initializes every physical row before this scan.
                var cells = _block[frame.RowIndex - _blockStart].Cells!;
                if (frame.Id == BIFFRECORDTYPE.MULRK)
                {
                    var lastColumn = frame.LastColumnIndex;
                    for (var column = (int)frame.ColumnIndex; column <= lastColumn; column++)
                        cells.Add(_decoder.Decode(_stream, frame, ixfe, column));
                }
                else
                {
                    cells.Add(_decoder.Decode(_stream, frame, ixfe, frame.ColumnIndex));
                }

                ixfe = null;
            }
            else if (frame.Id == BIFFRECORDTYPE.IXFE)
            {
                ixfe = frame.ReadUInt16(0);
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

            if (_currentRow >= 0 && (_sourceEnded || _pendingFrame.RowIndex != _currentRow))
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

            var frame = _pendingFrame;
            if (_currentRow < 0)
            {
                if (frame.RowIndex <= _lastCompletedRow)
                    throw new InvalidOperationException("File contains out-of-order cells. Use SinglePassMode = false to read this file.");

                _currentRow = frame.RowIndex;
            }

            if (frame.IsEmptyCell || (frame.Id == BIFFRECORDTYPE.MULRK && frame.LastColumnIndex < frame.ColumnIndex))
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

            var lastColumn = frame.Id == BIFFRECORDTYPE.MULRK ? frame.LastColumnIndex : frame.ColumnIndex;
            var cell = _decoder.Decode(_stream, frame, _ixfe, _pendingColumn++);
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
        while (_stream.ReadFrame(out var frame))
        {
            if (frame.IsCell)
            {
                _firstRecord = false;
                _pendingFrame = frame;
                _hasPendingRecord = true;
                _pendingColumn = frame.ColumnIndex;
                return;
            }

            if (frame.Id == BIFFRECORDTYPE.EOF || (frame.Id == BIFFRECORDTYPE.BOF && !_firstRecord))
            {
                _sourceEnded = true;
                return;
            }

            _firstRecord = false;
            if (frame.Id is BIFFRECORDTYPE.ROW or BIFFRECORDTYPE.ROW_V2)
            {
                double rowHeight;
                if (frame.Id == BIFFRECORDTYPE.ROW_V2)
                {
                    var heightBits = frame.ReadUInt16(6);
                    bool useDefault = (heightBits & 0x8000) != 0;
                    rowHeight = (useDefault ? _worksheet.DefaultRowHeight : heightBits & 0x7FFFF) / 20.0;
                }
                else
                {
                    var heightBits = frame.ReadUInt16(6);
                    var flags = frame.ReadUInt16(12);
                    bool useDefault = (heightBits & 0x8000) != 0;
                    int actualHeight = (flags & 0x20) == 0 ? heightBits & 0x7FFFF : 0;
                    rowHeight = (useDefault ? _worksheet.DefaultRowHeight : actualHeight) / 20.0;
                }

                _rowHeights[frame.RowIndex] = rowHeight;
            }
            else if (frame.Id == BIFFRECORDTYPE.IXFE)
            {
                _ixfe = frame.ReadUInt16(0);
            }
        }

        _sourceEnded = true;
    }

    private void ReleasePendingRecord()
    {
        if (!_hasPendingRecord)
            return;

        _pendingFrame = default;
        _hasPendingRecord = false;
    }

    private struct BufferedRow
    {
        public bool Present;
        public double Height;
        public List<Cell>? Cells;
    }
}