using System.Text;

namespace ExcelDataReader.Core.CsvFormat;

/// <summary>
/// Low level, reentrant CSV parser. Call ParseBuffer() in a loop, and finally Flush() to empty the internal buffers.
/// </summary>
internal sealed class CsvParser
{
    public CsvParser(char separator, Encoding encoding, char? quoteChar = null, bool trimWhiteSpace = true)
    {
        Separator = separator;
        QuoteChar = quoteChar;

        Decoder = encoding.GetDecoder();
        Decoder.Fallback = new DecoderExceptionFallback();

        var bufferSize = 1024;
        CharBuffer = new char[bufferSize];

        State = CsvState.PreValue;

        TrimWhiteSpace = trimWhiteSpace;
    }

    private enum CsvState
    {
        PreValue,
        Value,
        QuotedValue,
        QuotedValueQuote,
        Separator,
        Linebreak,
        EndOfFile,
    }

    private CsvState State { get; set; }

    private char? QuoteChar { get; }

    private int TrailingWhitespaceCount { get; set; }

    private Decoder Decoder { get; }

    private bool HasCarriageReturn { get; set; }

    private char Separator { get; }

    private char[] CharBuffer { get; set; }

    private StringBuilder ValueResult { get; set; } = new();

    private List<string> RowResult { get; set; } = [];

    private List<List<string>> RowsResult { get; set; } = [];

    private bool TrimWhiteSpace { get; }

    public void ParseBuffer(byte[] bytes, int offset, int count, out List<List<string>> rows)
    {
        while (count > 0)
        {
            Decoder.Convert(bytes, offset, count, CharBuffer, 0, CharBuffer.Length, false, out var bytesUsed, out var charsUsed, out _);

            offset += bytesUsed;
            count -= bytesUsed;

            for (var i = 0; i < charsUsed; i++)
            {
                ParseChar(CharBuffer[i], 1);
            }
        }

        rows = RowsResult;
        RowsResult = [];
    }

    public void Flush(out List<List<string>> rows)
    {
        if (ValueResult.Length > 0 || RowResult.Count > 0)
        {
            AddValueToRow();
            AddRowToResult();
        }

        rows = RowsResult;
        RowsResult = [];
    }

    private void ParseChar(char c, int bytesUsed)
    {
        var parsed = false;
        while (!parsed)
        {
            parsed = State switch
            {
                CsvState.PreValue => ReadPreValue(c, bytesUsed),
                CsvState.Value => ReadValue(c, bytesUsed),
                CsvState.QuotedValue => ReadQuotedValue(c),
                CsvState.QuotedValueQuote => ReadQuotedValueQuote(c),
                CsvState.Separator => ReadSeparator(),
                CsvState.Linebreak => ReadLinebreak(c),
                _ => throw new InvalidOperationException("Unhandled parser state: " + State)
            };
        }
    }

    private bool ReadPreValue(char c, int bytesUsed)
    {
        if (IsWhitespace(c) && TrimWhiteSpace)
        {
            return true;
        }
        else if (QuoteChar.HasValue && c == QuoteChar.Value)
        {
            State = CsvState.QuotedValue;
            return true;
        }
        else if (c == Separator)
        {
            State = CsvState.Separator;
            return false;
        }
        else if (c == '\r' || c == '\n' || (c == '\0' && bytesUsed == 0))
        {
            State = CsvState.Linebreak;
            return false;
        }
        else if (c == '\0' && bytesUsed == 0)
        {
            State = CsvState.EndOfFile;
            return false;
        }
        else
        {
            State = CsvState.Value;
            return false;
        }
    }

    private bool ReadValue(char c, int bytesUsed)
    {
        if (c == Separator)
        {
            State = CsvState.Separator;
            return false;
        }
        else if (c == '\r' || c == '\n')
        {
            State = CsvState.Linebreak;
            return false;
        }
        else if (c == '\0' && bytesUsed == 0)
        {
            State = CsvState.EndOfFile;
            return false;
        }
        else
        {
            if (IsWhitespace(c) && TrimWhiteSpace)
            {
                TrailingWhitespaceCount++;
            }
            else
            {
                TrailingWhitespaceCount = 0;
            }

            ValueResult.Append(c);
            return true;
        }
    }

    private bool ReadQuotedValue(char c)
    {
        if (c == QuoteChar)
        {
            State = CsvState.QuotedValueQuote;
        }
        else
        {
            ValueResult.Append(c);
        }

        return true;
    }

    private bool ReadQuotedValueQuote(char c)
    {
        if (c == QuoteChar)
        {
            // Is escaped quote
            ValueResult.Append(c);
            State = CsvState.QuotedValue;
            return true;
        }
        else
        {
            // End of quote, read remainder of field as a regular value until separator
            State = CsvState.Value;
            return false;
        }
    }

    private bool ReadSeparator()
    {
        AddValueToRow();
        State = CsvState.PreValue;
        return true;
    }

    private bool ReadLinebreak(char c)
    {
        if (HasCarriageReturn)
        {
            HasCarriageReturn = false;
            AddValueToRow();
            AddRowToResult();
            State = CsvState.PreValue;
            return c == '\n';
        }
        else if (c == '\r')
        {
            HasCarriageReturn = true;
            return true;
        }
        else
        {
            AddValueToRow();
            AddRowToResult();
            State = CsvState.PreValue;
            return true;
        }
    }

    private void AddValueToRow()
    {
        RowResult.Add(ValueResult.ToString(0, ValueResult.Length - TrailingWhitespaceCount));
        ValueResult = new StringBuilder();
        TrailingWhitespaceCount = 0;
    }

    private void AddRowToResult()
    {
        RowsResult.Add(RowResult);
        RowResult = [];
    }

    private bool IsWhitespace(char c)
    {
        if (c == ' ')
        {
            return true;
        }

        if (Separator != '\t' && c == '\t')
        {
            return true;
        }

        return false;
    }
}