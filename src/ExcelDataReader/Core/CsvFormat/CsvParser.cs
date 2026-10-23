using System.Text;

namespace ExcelDataReader.Core.CsvFormat;

/// <summary>
/// Reentrant CSV character parser. The byte offsets and decoded character cursor survive callback pauses.
/// Scan mode tracks field lengths using the same grammar without building field values.
/// </summary>
internal sealed class CsvParser
{
    private readonly char _separator;
    private readonly char? _quoteChar;
    private readonly char? _escapeChar;
    private readonly bool _trimWhiteSpace;
    private readonly Decoder _decoder;
    private readonly char[] _chars = new char[1024];
    private readonly StringBuilder? _value;
    private CsvState _state;
    private PendingEvent _pending;
    private int _charOffset;
    private int _charCount;
    private int _valueLength;
    private int _trailingWhitespace;
    private int _fieldCount;
    private int _rowIndex;
    private int _lastFieldLength;
    private bool _rowStarted;
    private bool _carriageReturn;
    private bool _endRow;
    private bool _flushing;
    private bool _finished;

    public CsvParser(char separator, Encoding encoding, char? quoteChar = null, bool trimWhiteSpace = true, char? escapeChar = null, WorksheetReadMode mode = WorksheetReadMode.Materialize)
    {
        _separator = separator;
        _quoteChar = quoteChar;
        _escapeChar = escapeChar;
        _trimWhiteSpace = trimWhiteSpace;
        _decoder = encoding.GetDecoder();
        _decoder.Fallback = new DecoderExceptionFallback();
        if (mode == WorksheetReadMode.Materialize)
            _value = new StringBuilder();
    }

    private enum CsvState
    {
        PreValue,
        Value,
        QuotedValue,
        QuotedValueEscape,
        QuotedValueQuote,
        Linebreak,
    }

    private enum PendingEvent
    {
        None,
        Row,
        Field,
        RowEnd,
    }

    public WorksheetReadStatus ParseBuffer<THandler>(byte[] bytes, ref int offset, ref int count, ref THandler handler)
        where THandler : struct, ICsvHandler
    {
        while (true)
        {
            if (EmitPending(ref handler) == WorksheetReadControl.Pause)
                return WorksheetReadStatus.Paused;

            if (_charOffset < _charCount)
            {
                if (ParseChar(_chars[_charOffset]))
                    _charOffset++;
                continue;
            }

            if (count == 0)
                return WorksheetReadStatus.EndOfInput;

            _decoder.Convert(bytes, offset, count, _chars, 0, _chars.Length, false, out var bytesUsed, out _charCount, out _);
            offset += bytesUsed;
            count -= bytesUsed;
            _charOffset = 0;
        }
    }

    public WorksheetReadStatus Flush<THandler>(ref THandler handler)
        where THandler : struct, ICsvHandler
    {
        if (!_flushing)
        {
            _flushing = true;

            // Match the original EOF rule: an empty final field alone does not create a row.
            if (_valueLength > 0 || _fieldCount > 0)
                QueueField(endRow: true);
            else
                _finished = true;
        }

        if (EmitPending(ref handler) == WorksheetReadControl.Pause)
            return WorksheetReadStatus.Paused;

        _finished = true;
        return WorksheetReadStatus.EndOfInput;
    }

    private WorksheetReadControl EmitPending<THandler>(ref THandler handler)
        where THandler : struct, ICsvHandler
    {
        while (_pending != PendingEvent.None)
        {
            switch (_pending)
            {
                case PendingEvent.Row:
                    _rowStarted = true;
                    _pending = PendingEvent.Field;
                    if (handler.OnRow(_rowIndex) == WorksheetReadControl.Pause)
                        return WorksheetReadControl.Pause;
                    break;
                case PendingEvent.Field:
                    _lastFieldLength = _valueLength - _trailingWhitespace;
                    string? value = _value?.ToString(0, _lastFieldLength);
                    _value?.Clear();
                    _valueLength = 0;
                    _trailingWhitespace = 0;
                    var column = _fieldCount++;
                    _pending = _endRow ? PendingEvent.RowEnd : PendingEvent.None;
                    if (value != null && handler.OnField(column, value) == WorksheetReadControl.Pause)
                        return WorksheetReadControl.Pause;
                    break;
                case PendingEvent.RowEnd:
                    int fields = _fieldCount;
                    bool empty = fields == 1 && _lastFieldLength == 0;
                    _fieldCount = 0;
                    _rowStarted = false;
                    _rowIndex++;
                    _pending = PendingEvent.None;
                    if (handler.OnRowEnd(fields, empty, _flushing) == WorksheetReadControl.Pause)
                        return WorksheetReadControl.Pause;
                    break;
            }
        }

        return WorksheetReadControl.Continue;
    }

    private bool ParseChar(char c)
    {
        switch (_state)
        {
            case CsvState.PreValue:
                if (_trimWhiteSpace && IsWhitespace(c))
                    return true;
                if (c == _quoteChar)
                {
                    _state = CsvState.QuotedValue;
                    return true;
                }

                _state = CsvState.Value;
                return false;
            case CsvState.Value:
                if (c == _separator)
                {
                    _state = CsvState.PreValue;
                    QueueField(endRow: false);
                }
                else if (c is '\r' or '\n')
                {
                    _state = CsvState.Linebreak;
                    return false;
                }
                else
                {
                    _trailingWhitespace = _trimWhiteSpace && IsWhitespace(c) ? _trailingWhitespace + 1 : 0;
                    Append(c);
                }

                return true;
            case CsvState.QuotedValue:
                if (c == _quoteChar)
                    _state = CsvState.QuotedValueQuote;
                else if (c == _escapeChar)
                    _state = CsvState.QuotedValueEscape;
                else
                    Append(c);
                return true;
            case CsvState.QuotedValueEscape:
                Append(c);
                _state = CsvState.QuotedValue;
                return true;
            case CsvState.QuotedValueQuote:
                if (c == _quoteChar)
                {
                    Append(c);
                    _state = CsvState.QuotedValue;
                    return true;
                }

                _state = CsvState.Value;
                return false;
            case CsvState.Linebreak:
                if (_carriageReturn)
                {
                    _carriageReturn = false;
                    _state = CsvState.PreValue;
                    QueueField(endRow: true);
                    return c == '\n';
                }

                if (c == '\r')
                {
                    _carriageReturn = true;
                }
                else
                {
                    _state = CsvState.PreValue;
                    QueueField(endRow: true);
                }

                return true;
            default:
                throw new InvalidOperationException("Unhandled CSV parser state.");
        }
    }

    private void QueueField(bool endRow)
    {
        if (_finished)
            return;
        _endRow = endRow;
        _pending = _rowStarted ? PendingEvent.Field : PendingEvent.Row;
    }

    private void Append(char c)
    {
        _valueLength++;
        _value?.Append(c);
    }

    private bool IsWhitespace(char c) => c == ' ' || (_separator != '\t' && c == '\t');
}