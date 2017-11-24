using System;

namespace ExcelDataReader.Core.NumberFormat
{
    internal class Tokenizer
    {
        private string formatString;
        private int formatStringPosition = 0;

        public Tokenizer(string fmt)
        {
            formatString = fmt;
        }

        public int Position => formatStringPosition;

        public int Length => formatString.Length;

        public string Substring(int startIndex, int length)
        {
            return formatString.Substring(startIndex, length);
        }

        public int Peek(int offset = 0)
        {
            if (formatStringPosition + offset >= formatString.Length)
                return -1;
            return formatString[formatStringPosition + offset];
        }

        public int PeekUntil(int startOffset, int until)
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

        public bool PeekOneOf(int offset, string s)
        {
            foreach (var c in s)
            {
                if (Peek(offset) == c)
                {
                    return true;
                }
            }

            return false;
        }

        public void Advance(int characters = 1)
        {
            formatStringPosition = Math.Min(formatStringPosition + characters, formatString.Length);
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
            if (formatStringPosition + s.Length > formatString.Length)
                return false;

            for (var i = 0; i < s.Length; i++)
            {
                var c1 = s[i];
                var c2 = (char)Peek(i);
                if (ignoreCase)
                {
                    if (char.ToLower(c1) != char.ToLower(c2))
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
    }
}
