using System.Data;
using System.Diagnostics.CodeAnalysis;
using ExcelDataReader.Core;

namespace ExcelDataReader;

/// <summary>
/// A generic implementation of IExcelDataReader using worksheet push consumption.
/// </summary>
/// <typeparam name="TWorkbook">A type implementing IWorkbook.</typeparam>
/// <typeparam name="TWorksheet">A type implementing IWorksheet.</typeparam>
internal abstract class ExcelDataReader<TWorkbook, TWorksheet> : IExcelDataReader
    where TWorkbook : IWorkbook<TWorksheet>
    where TWorksheet : IWorksheet
{
    private IEnumerator<TWorksheet>? _worksheetIterator;
    private IWorksheetReader? _rowReader;
    private IEnumerator<TWorksheet>? _cachedWorksheetIterator;
    private List<TWorksheet>? _cachedWorksheets;
    private int _idx;
    private bool _singlePassMode;
    private int _nextRowIndex;
    private int _pendingRowIndex = -1;
    private double _pendingRowHeight;
    private double _rowHeight;
    private bool _rowReady;
    private Dictionary<long, CellCommentThread>? _comments;
    private bool _commentsLoaded;

    ~ExcelDataReader()
    {
        Dispose(false);
    }

    public string? Name => _worksheetIterator?.Current?.Name;

    public string? CodeName => _worksheetIterator?.Current?.CodeName;

    public string? VisibleState => _worksheetIterator?.Current?.VisibleState;

    public virtual string? CsvSeparator => null;

    public int ActiveSheet => GetWorkbook().ActiveSheet;

    public bool IsActiveSheet => _idx == GetWorkbook().ActiveSheet;

    public HeaderFooter? HeaderFooter => _worksheetIterator?.Current?.HeaderFooter;

    // We shouldn't expose the internal array here. 
    public CellRange[] MergeCells => _worksheetIterator?.Current?.MergeCells ?? [];

    public int Depth => 0;

    public int ResultsCount => Workbook?.ResultsCount ?? -1;

    public bool IsClosed { get; private set; }

    public int FieldCount => _singlePassMode
        ? Math.Max(_worksheetIterator?.Current?.FieldCount ?? 0, RowCells?.Length ?? 0)
        : (_worksheetIterator?.Current?.FieldCount ?? 0);

    public int RowCount => _singlePassMode
        ? throw new InvalidOperationException("RowCount is not available in SinglePassMode.")
        : (_worksheetIterator?.Current?.RowCount ?? 0);

    public CellRange? Dimension => _worksheetIterator?.Current?.Dimension;

    public int RecordsAffected => throw new NotSupportedException();

    public double RowHeight => _rowHeight;

    protected bool SinglePassMode { set => _singlePassMode = value; }

    protected TWorkbook? Workbook { get; set; }

    protected TWorksheet? CurrentWorksheet => _worksheetIterator == null ? default : _worksheetIterator.Current;

    private Cell?[]? RowCells { get; set; }

    public object this[int i] => GetValue(i);

    public object this[string name] => throw new NotSupportedException();

    public bool GetBoolean(int i) => (bool)GetValue(i);

    public byte GetByte(int i) => (byte)GetValue(i);

    public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length)
        => throw new NotSupportedException();

    public char GetChar(int i) => (char)GetValue(i);

    public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length)
         => throw new NotSupportedException();

    public IDataReader GetData(int i) => throw new NotSupportedException();

    public string GetDataTypeName(int i) => throw new NotSupportedException();

    public DateTime GetDateTime(int i) => (DateTime)GetValue(i);

    public decimal GetDecimal(int i) => (decimal)GetValue(i);

    public double GetDouble(int i) => (double)GetValue(i);

#if NET8_0_OR_GREATER
    [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
#endif
    public Type GetFieldType(int i) => GetValue(i)?.GetType() ?? typeof(DBNull);

    public float GetFloat(int i) => (float)GetValue(i);

    public Guid GetGuid(int i) => (Guid)GetValue(i);

    public short GetInt16(int i) => (short)GetValue(i);

    public int GetInt32(int i) => (int)GetValue(i);

    public long GetInt64(int i) => (long)GetValue(i);

    public string GetName(int i) => throw new NotSupportedException();

    public int GetOrdinal(string name) => throw new NotSupportedException();

    /// <inheritdoc />
    public DataTable GetSchemaTable() => throw new NotSupportedException();

    public string GetString(int i) => (string)GetValue(i);

    public object GetValue(int i)
    {
        if (RowCells == null)
            throw new InvalidOperationException("No data exists for the row/column.");

        return RowCells[i]?.Value ?? DBNull.Value;
    }

    public int GetValues(object[] values)
    {
        if (RowCells == null)
            throw new InvalidOperationException("No data exists for the row/column.");

        int readingLenth = values.Length > FieldCount ? FieldCount : values.Length;
        for (int i = 0; i < readingLenth; i++)
        {
            values[i] = RowCells[i]?.Value ?? DBNull.Value;
        }

        return readingLenth;
    }

    public bool IsDBNull(int i) => GetValue(i) is DBNull;

    public string? GetNumberFormatString(int i)
    {
        if (RowCells == null)
            throw new InvalidOperationException("No data exists for the row/column.");
        if (RowCells[i]?.EffectiveStyle is not { } effectiveStyle)
            return null;
        return GetWorkbook().GetNumberFormatString(effectiveStyle.NumberFormatIndex, null)?.FormatString;
    }

    public string? GetNumberFormatString(int i, IFormatProvider? provider)
    {
        if (RowCells == null)
            throw new InvalidOperationException("No data exists for the row/column.");
        if (RowCells[i]?.EffectiveStyle is not { } effectiveStyle)
            return null;
        return GetWorkbook().GetNumberFormatString(effectiveStyle.NumberFormatIndex, provider)?.FormatString;
    }

    public int GetNumberFormatIndex(int i)
    {
        if (RowCells == null)
            throw new InvalidOperationException("No data exists for the row/column.");
        if (RowCells[i]?.EffectiveStyle is not { } effectiveStyle)
            return -1;
        return effectiveStyle.NumberFormatIndex;
    }

    public double GetColumnWidth(int i)
    {
        const double DefaultColumnWidth = 8.43D;

        if (i >= FieldCount)
        {
            throw new ArgumentException($"Column at index {i} does not exist.", nameof(i));
        }

        var columnWidths = _worksheetIterator?.Current?.ColumnWidths;
        if (columnWidths != null)
        {
            foreach (var columnWidth in columnWidths)
            {
                if (i >= columnWidth.Minimum && i <= columnWidth.Maximum)
                {
                    return columnWidth.Hidden ? 0 : columnWidth.Width ?? DefaultColumnWidth;
                }
            }
        }

        return DefaultColumnWidth;
    }

    public CellStyle GetCellStyle(int i)
    {
        if (RowCells == null)
            throw new InvalidOperationException("No data exists for the row/column.");

        return (RowCells[i]?.EffectiveStyle ?? ExtendedFormat.Zero).CellStyle;
    }

    public CellError? GetCellError(int i)
    {
        if (RowCells == null)
            throw new InvalidOperationException("No data exists for the row/column.");

        return RowCells[i]?.Error;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201:Do not raise reserved exception types", Justification = "Match existing cell accessor column validation.")]
    public CellCommentThread? GetComment(int i)
    {
        if (RowCells == null || _rowReader == null)
            throw new InvalidOperationException("No data exists for the row/column.");
        if ((uint)i >= (uint)FieldCount)
            throw new IndexOutOfRangeException();
        if (!_commentsLoaded)
        {
            _comments = _worksheetIterator!.Current.LoadComments();
            _commentsLoaded = true;
        }

        return _comments != null && _comments.TryGetValue(((long)(_nextRowIndex - 1) << 32) | (uint)i, out var comment) ? comment : null;
    }

    /// <inheritdoc />
    public void Reset()
    {
        _rowReader?.Dispose();
        _worksheetIterator?.Dispose();

        _worksheetIterator = null;
        _rowReader = null;

        _idx = 0;

        ResetSheetData();

        if (Workbook != null)
        {
            _worksheetIterator = ReadWorksheetsWithCache().GetEnumerator();
            if (!_worksheetIterator.MoveNext())
            {
                _worksheetIterator.Dispose();
                _worksheetIterator = null;
                return;
            }

            _rowReader = _worksheetIterator.Current.OpenReader();
        }
    }

    /// <inheritdoc />
    public void ResetWorksheet()
    {
        if (IsClosed || _worksheetIterator == null)
            throw new InvalidOperationException("No worksheet is selected.");

        var reader = _rowReader;
        _rowReader = null;
        ResetSheetData();
        reader?.Dispose();
        _rowReader = _worksheetIterator.Current.OpenReader();
    }

    public virtual void Close()
    {
        if (IsClosed)
            return;

        _rowReader?.Dispose();
        _worksheetIterator?.Dispose();
        _cachedWorksheetIterator?.Dispose();

        _worksheetIterator = null;
        _rowReader = null;
        _cachedWorksheetIterator = null;
        _cachedWorksheets = null;
        ResetSheetData();
        try
        {
            Workbook?.Dispose();
        }
        finally
        {
            Workbook = default;
            IsClosed = true;
        }
    }

    public bool NextResult()
    {
        if (_worksheetIterator == null)
        {
            return false;
        }

        ResetSheetData();

        _rowReader?.Dispose();
        _rowReader = null;

        if (!_worksheetIterator.MoveNext())
        {
            _worksheetIterator.Dispose();
            _worksheetIterator = null;
            return false;
        }

        _rowReader = _worksheetIterator.Current.OpenReader();

        _idx++;

        return true;
    }

    public bool Read()
    {
        if (_worksheetIterator == null || _rowReader == null)
        {
            return false;
        }

        if (_pendingRowIndex >= 0)
        {
            ClearRow();
            if (_pendingRowIndex > _nextRowIndex)
                return ReadGapRow();
            _rowHeight = _pendingRowHeight;
        }

        _rowReady = false;
        var handler = new RowHandler(this);
        WorksheetReadStatus status;
        try
        {
            status = _rowReader.Consume(ref handler);
        }
        catch (Exception exception)
        {
            var failedReader = _rowReader;
            _rowReader = null;
            ResourceCleanup.DisposeAll(exception, failedReader);
            throw;
        }

        if (_rowReady)
            return true;
        if (_pendingRowIndex > _nextRowIndex)
            return ReadGapRow();
        if (status != WorksheetReadStatus.EndOfInput)
            throw new InvalidOperationException("Worksheet paused without completing a row.");

        _rowReader.Dispose();
        _rowReader = null;
        _rowHeight = 0;
        return false;
    }

    public void Dispose()
    {
        Dispose(true);

        GC.SuppressFinalize(this);
    }

    private void Dispose(bool disposing)
    {
        if (disposing)
            Close();
    }

    private IEnumerable<TWorksheet> ReadWorksheetsWithCache()
    {
        // Iterate TWorkbook.ReadWorksheets() only once and cache the 
        // worksheet instances, which are expensive to create. 
        if (_cachedWorksheets != null)
        {
            foreach (var worksheet in _cachedWorksheets)
            {
                yield return worksheet;
            }

            if (_cachedWorksheetIterator == null)
            {
                yield break;
            }
        }
        else
        {
            _cachedWorksheets = [];
        }

        if (Workbook is null)
        {
            yield break;
        }

        _cachedWorksheetIterator ??= Workbook.ReadWorksheets().GetEnumerator();

        while (_cachedWorksheetIterator.MoveNext())
        {
            _cachedWorksheets.Add(_cachedWorksheetIterator.Current);
            yield return _cachedWorksheetIterator.Current;
        }

        _cachedWorksheetIterator.Dispose();
        _cachedWorksheetIterator = null;
    }

    private void ResetSheetData()
    {
        _comments = null;
        _commentsLoaded = false;
        RowCells = null;
        _nextRowIndex = 0;
        _pendingRowIndex = -1;
        _rowHeight = 0;
        _rowReady = false;
    }

    private void ClearRow()
    {
        RowCells ??= new Cell?[FieldCount];
        Array.Clear(RowCells, 0, RowCells.Length);
    }

    private bool ReadGapRow()
    {
        _rowHeight = _worksheetIterator!.Current.DefaultRowHeight;
        _nextRowIndex++;
        return true;
    }

    private TWorkbook GetWorkbook()
    {
        return Workbook ?? throw new InvalidOperationException("Workbook is not initialized.");
    }

    private readonly struct RowHandler(ExcelDataReader<TWorkbook, TWorksheet> owner) : IWorksheetHandler
    {
        public WorksheetReadControl OnDimension(CellRange range) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnCodeName(string? codeName) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnDefaultRowHeight(double? height) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnColumn(Column column) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnMergeCell(CellRange range) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnHeaderFooter(HeaderFooter headerFooter) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnSheetDataBegin() => WorksheetReadControl.Continue;

        public WorksheetReadControl OnScanResult(int maxRowIndex, int maxColumnIndex) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnRow(int index, bool hidden, double? height)
        {
            if (index < owner._nextRowIndex)
                throw new InvalidOperationException("Worksheet rows must be in ascending index order.");

            if (owner._pendingRowIndex == index)
            {
                owner._rowHeight = hidden ? 0 : height ?? owner._worksheetIterator!.Current.DefaultRowHeight;
                return WorksheetReadControl.Continue;
            }

            owner._pendingRowIndex = index;
            owner._pendingRowHeight = hidden ? 0 : height ?? owner._worksheetIterator!.Current.DefaultRowHeight;
            owner.ClearRow();
            if (index > owner._nextRowIndex)
                return WorksheetReadControl.Pause;

            owner._rowHeight = owner._pendingRowHeight;
            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnCell(in Cell cell)
        {
            var cells = owner.RowCells!;
            if (cell.ColumnIndex >= cells.Length)
            {
                if (!owner._singlePassMode)
                    return WorksheetReadControl.Continue;

                Array.Resize(ref cells, cell.ColumnIndex + 1);
                owner.RowCells = cells;
            }

            cells[cell.ColumnIndex] = cell;
            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnRowEnd()
        {
            owner._nextRowIndex = owner._pendingRowIndex + 1;
            owner._pendingRowIndex = -1;
            owner._rowReady = true;
            return WorksheetReadControl.Pause;
        }
    }
}