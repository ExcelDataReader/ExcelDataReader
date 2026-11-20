using System.Globalization;

namespace ExcelDataReader.Core.NumberFormat;

internal sealed class Tokenizer(string fmt)
{
    public int Position { get; private set; }

    public int Length => Source.Length;

    public string Source { get; } = fmt;

    public double ParseDouble(int startIndex, int length)
    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        return double.Parse(Source.AsSpan(startIndex, length), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
#else
        return double.Parse(Source.Substring(startIndex, length), CultureInfo.InvariantCulture);
#endif
    }

    public int ParseInt32(List<Token> tokens, int first, int count)
    {
        int start = tokens[first].Start;
        int length = tokens[first + count - 1].Start + 1 - start;
        if (length == count)
        {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
            return int.Parse(Source.AsSpan(start, length), NumberStyles.Integer, CultureInfo.InvariantCulture);
#else
            return int.Parse(Source.Substring(start, length), CultureInfo.InvariantCulture);
#endif
        }

        // Directives can separate denominator digits. The first digit is nonzero.
        if (count > 10)
            throw new OverflowException();
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        Span<char> digits = stackalloc char[10];
        for (int i = 0; i < count; i++)
            digits[i] = Source[tokens[first + i].Start];
        return int.Parse(digits[..count], NumberStyles.Integer, CultureInfo.InvariantCulture);
#else
        var digits = new char[count];
        for (int i = 0; i < count; i++)
            digits[i] = Source[tokens[first + i].Start];
        return int.Parse(new string(digits), CultureInfo.InvariantCulture);
#endif
    }

    public int Peek(int offset = 0)
    {
        if (Position + offset >= Source.Length)
            return -1;
        return Source[Position + offset];
    }

    public void Advance(int characters = 1)
    {
        Position = Math.Min(Position + characters, Source.Length);
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
        if (Position + s.Length > Source.Length)
            return false;

        for (var i = 0; i < s.Length; i++)
        {
            var c1 = s[i];
            var c2 = (char)Peek(i);
            if (ignoreCase)
            {
                if (!Helpers.EqualsAsciiIgnoreCase(c1, c2))
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

    private int PeekUntil(int startOffset, char until)
    {
        int start = Position + startOffset;
        int end = Source.IndexOf(until, start);
        return end < 0 ? 0 : end - start + 1;
    }

    private bool PeekOneOf(int offset, string s)
    {
        int c = Peek(offset);
        return c >= 0 && s.Contains((char)c);
    }
}