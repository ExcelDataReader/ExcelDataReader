using System.Globalization;

namespace ExcelDataReader.Core.NumberFormat;

internal sealed class Tokenizer(string fmt)
{
    private readonly string _formatString = fmt;

    public int Position { get; private set; }

    public int Length => _formatString.Length;

    public string Source => _formatString;

    public double ParseDouble(int startIndex, int length)
    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        return double.Parse(_formatString.AsSpan(startIndex, length), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
#else
        return double.Parse(_formatString.Substring(startIndex, length), CultureInfo.InvariantCulture);
#endif
    }

    public int ParseInt32(List<Token> tokens, int first, int count)
    {
        int start = tokens[first].Start;
        int length = tokens[first + count - 1].Start + 1 - start;
        if (length == count)
        {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
            return int.Parse(_formatString.AsSpan(start, length), NumberStyles.Integer, CultureInfo.InvariantCulture);
#else
            return int.Parse(_formatString.Substring(start, length), CultureInfo.InvariantCulture);
#endif
        }

        // Directives can separate denominator digits. The first digit is nonzero.
        if (count > 10)
            throw new OverflowException();
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        Span<char> digits = stackalloc char[10];
        for (int i = 0; i < count; i++)
            digits[i] = _formatString[tokens[first + i].Start];
        return int.Parse(digits[..count], NumberStyles.Integer, CultureInfo.InvariantCulture);
#else
        var digits = new char[count];
        for (int i = 0; i < count; i++)
            digits[i] = _formatString[tokens[first + i].Start];
        return int.Parse(new string(digits), CultureInfo.InvariantCulture);
#endif
    }

    public int Peek(int offset = 0)
    {
        if (Position + offset >= _formatString.Length)
            return -1;
        return _formatString[Position + offset];
    }

    public void Advance(int characters = 1)
    {
        Position = Math.Min(Position + characters, _formatString.Length);
    }

    public bool ReadOneOrMore(int c)
    {
        if (Peek() != c)
            return false;

        while (Peek() == c)
            Advance();

        return true;
    }

    public bool ReadOneOf(string s)
    {
        if (PeekOneOf(0, s))
        {
            Advance();
            return true;
        }

        return false;
    }

    public bool ReadString(string s, bool ignoreCase = false)
    {
        if (Position + s.Length > _formatString.Length)
            return false;

        for (var i = 0; i < s.Length; i++)
        {
            var c1 = s[i];
            var c2 = (char)Peek(i);
            if (ignoreCase)
            {
                if (char.ToLower(c1, CultureInfo.InvariantCulture) != char.ToLower(c2, CultureInfo.InvariantCulture))
                    return false;
            }
            else
            {
                if (c1 != c2)
                    return false;
            }
        }

        Advance(s.Length);
        return true;
    }

    public bool ReadEnclosed(char open, char close)
    {
        if (Peek() == open)
        {
            int length = PeekUntil(1, close);
            if (length > 0)
            {
                Advance(1 + length);
                return true;
            }
        }

        return false;
    }
    
    private int PeekUntil(int startOffset, int until)
    {
        int offset = startOffset;
        while (true)
        {
            var c = Peek(offset++);
            if (c == -1)
                break;
            if (c == until)
                return offset - startOffset;
        }

        return 0;
    }

    private bool PeekOneOf(int offset, string s)
    {
        int c = Peek(offset);
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        return c >= 0 && s.Contains((char)c);
#else
        return c >= 0 && s.IndexOf((char)c) >= 0;
#endif
    }
}
