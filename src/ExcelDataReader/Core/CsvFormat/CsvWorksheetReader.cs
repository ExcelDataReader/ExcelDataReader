namespace ExcelDataReader.Core.CsvFormat;

internal sealed class CsvWorksheetReader : IWorksheetReader
{
    private readonly Stream _stream;
    private readonly CsvParser _parser;
    private readonly byte[] _bytes = new byte[1024];
    private int _offset;
    private int _count;
    private bool _begun;
    private bool _disposed;

    public CsvWorksheetReader(CsvWorksheet worksheet, int bomLength)
    {
        _stream = worksheet.Stream;
        _stream.Seek(bomLength, SeekOrigin.Begin);
        _parser = new CsvParser(worksheet.Separator, worksheet.Encoding, worksheet.QuoteChar, worksheet.TrimWhiteSpace, worksheet.EscapeChar);
    }

    public WorksheetReadStatus Consume<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler
    {
        if (_disposed)
            return WorksheetReadStatus.EndOfInput;

        if (!_begun)
        {
            _begun = true;
            if (handler.OnSheetDataBegin() == WorksheetReadControl.Pause)
                return WorksheetReadStatus.Paused;
        }

        var adapter = new Handler<THandler>(handler);
        try
        {
            while (true)
            {
                if (_parser.ParseBuffer(_bytes, ref _offset, ref _count, ref adapter) == WorksheetReadStatus.Paused)
                    return WorksheetReadStatus.Paused;

                if (_stream.Position >= _stream.Length)
                    return _parser.Flush(ref adapter);

                _offset = 0;
                _count = _stream.ReadAtLeast(_bytes, 0, _bytes.Length);
            }
        }
        finally
        {
            handler = adapter.Target;
        }
    }

    public void Dispose() => _disposed = true;

    private struct Handler<THandler>(THandler target) : ICsvHandler
        where THandler : struct, IWorksheetHandler
    {
        public THandler Target = target;

        public WorksheetReadControl OnRow(int index) => Target.OnRow(index, false, null);

        public WorksheetReadControl OnField(int index, string value)
        {
            var cell = new Cell(index, value, ExtendedFormat.Zero, null);
            return Target.OnCell(in cell);
        }

        public WorksheetReadControl OnRowEnd(int fieldCount, bool empty, bool flushing) => Target.OnRowEnd();
    }
}