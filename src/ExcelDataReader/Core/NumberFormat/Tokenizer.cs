﻿using System.Globalization;

namespace ExcelDataReader.Core.NumberFormat;

internal sealed class Tokenizer(string fmt)
{
    private readonly string _formatString = fmt;

    public int Position { get; private set; }

    public int Length => _formatString.Length;

    public string Substring(int startIndex, int length)
    {
        return _formatString.Substring(startIndex, length);
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

    private bool PeekOneOf(int offset, string s) => s.Any(c => Peek(offset) == c);
}
