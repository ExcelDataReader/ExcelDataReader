using System.Text;

namespace ExcelDataReader.Core.CsvFormat;

/// <summary>
/// Reentrant CSV character parser. The byte offsets and decoded character cursor survive callback pauses.
/// Scan mode tracks field lengths using the same grammar without building field values.
/// </summary>
internal sealed class CsvParser
{
    private readonly char _separator;
    private readonly string? _delimiter;
    private readonly int[]? _delimiterPrefixes;
    private readonly char? _quoteChar;
    private readonly char? _escapeChar;
    private readonly bool _trimWhiteSpace;
    private readonly bool _trimTabs;
    private readonly Decoder _decoder;
    private readonly char[] _chars = new char[1024];
    private readonly StringBuilder? _value;
    private int _delimiterOffset;
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

    public CsvParser(string separator, Encoding encoding, char? quoteChar = null, bool trimWhiteSpace = true, char? escapeChar = null, WorksheetReadMode mode = WorksheetReadMode.Materialize)
    {
        if (string.IsNullOrEmpty(separator))
            throw new ArgumentException("CSV separator candidates must not be null or empty.", nameof(separator));
        if (separator.Length > 1)
        {
            foreach (char c in separator)
            {
                if (c is '\r' or '\n' || c == quoteChar)
                    throw new ArgumentException("Multi-character CSV separators must not contain line breaks or the quote character.", nameof(separator));
            }

            _delimiter = separator;
            _delimiterPrefixes = new int[separator.Length];
            for (int i = 1, prefix = 0; i < separator.Length; i++)
            {
                while (prefix > 0 && separator[i] != separator[prefix])
                    prefix = _delimiterPrefixes[prefix - 1];
                if (separator[i] == separator[prefix])
                    prefix++;
                _delimiterPrefixes[i] = prefix;
            }
        }

        _separator = separator[0];
        _trimTabs = separator != "\t";
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

            if (_delimiter == null)
            {
                while (_charOffset < _charCount)
                {
                    if (ParseChar(_chars[_charOffset]))
                        _charOffset++;
                    if (EmitPending(ref handler) == WorksheetReadControl.Pause)
                        return WorksheetReadStatus.Paused;
                }
            }
            else
            {
                while (_charOffset < _charCount)
                {
                    if (ParseDelimitedChar(_chars[_charOffset]))
                        _charOffset++;
                    if (EmitPending(ref handler) == WorksheetReadControl.Pause)
                        return WorksheetReadStatus.Paused;
                }
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
            FlushDelimiterPrefix(_delimiterOffset);
            _delimiterOffset = 0;

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

    private void ParseRun(bool quoted)
    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        var remaining = _chars.AsSpan(_charOffset, _charCount - _charOffset);
        int end;
        if (!quoted)
        {
            end = remaining.IndexOfAny(_separator, '\r', '\n');
        }
        else
        {
            end = _quoteChar.HasValue
                ? _escapeChar.HasValue ? remaining.IndexOfAny(_quoteChar.Value, _escapeChar.Value) : remaining.IndexOf(_quoteChar.Value)
                : _escapeChar.HasValue ? remaining.IndexOf(_escapeChar.Value) : -1;
        }

        int length = end < 0 ? remaining.Length : end;
#else
        int end = _charOffset;
        while (end < _charCount)
        {
            char c = _chars[end];
            if (quoted ? c == _quoteChar || c == _escapeChar : c == _separator || c is '\r' or '\n')
                break;
            end++;
        }

        int length = end - _charOffset;
#endif
        if (!quoted)
        {
            if (_trimWhiteSpace)
            {
                int last = length - 1;
                while (last >= 0 && IsWhitespace(_chars[_charOffset + last]))
                    last--;
                _trailingWhitespace = last < 0 ? _trailingWhitespace + length : length - last - 1;
            }
            else
            {
                _trailingWhitespace = 0;
            }
        }

        _valueLength += length;
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        _value?.Append(remaining[..length]);
#else
        _value?.Append(_chars, _charOffset, length);
#endif

        // ParseBuffer advances past the final character after ParseChar returns.
        _charOffset += length - 1;
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
                    ParseRun(quoted: false);
                }

                return true;
            case CsvState.QuotedValue:
                if (c == _quoteChar)
                {
                    _state = CsvState.QuotedValueQuote;
                }
                else if (c == _escapeChar)
                {
                    _state = CsvState.QuotedValueEscape;
                }
                else
                {
                    // Dense escaping has no useful text runs; leave it on the scalar path.
                    if (_charOffset + 1 < _charCount &&
                        _chars[_charOffset + 1] != _quoteChar && _chars[_charOffset + 1] != _escapeChar)
                    {
                        ParseRun(quoted: true);
                        return true;
                    }

                    Append(c);
                }

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

    private bool ParseDelimitedChar(char c)
    {
        if (_state == CsvState.QuotedValueQuote && c != _quoteChar)
            _state = CsvState.Value;

        if (_state is not (CsvState.PreValue or CsvState.Value))
            return ParseChar(c);

        while (_delimiterOffset > 0 && c != _delimiter![_delimiterOffset])
        {
            int prefix = _delimiterPrefixes![_delimiterOffset - 1];
            FlushDelimiterPrefix(_delimiterOffset - prefix);
            _delimiterOffset = prefix;
        }

        if (c == _delimiter![_delimiterOffset])
        {
            if (++_delimiterOffset == _delimiter.Length)
            {
                _delimiterOffset = 0;
                _state = CsvState.PreValue;
                QueueField(endRow: false);
            }

            return true;
        }

        return ParseChar(c);
    }

    private void FlushDelimiterPrefix(int length)
    {
        // Only unmatched prefix characters become field text; a complete separator is never trimmed.
        for (int i = 0; i < length; i++)
        {
            char c = _delimiter![i];
            if (_state == CsvState.PreValue && _trimWhiteSpace && IsWhitespace(c))
                continue;
            _state = CsvState.Value;
            Append(c);
            _trailingWhitespace = _trimWhiteSpace && IsWhitespace(c) ? _trailingWhitespace + 1 : 0;
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

    private bool IsWhitespace(char c) => c == ' ' || (_trimTabs && c == '\t');
}