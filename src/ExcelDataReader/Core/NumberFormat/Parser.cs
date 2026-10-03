using System.Globalization;

namespace ExcelDataReader.Core.NumberFormat;

internal static class Parser
{
    public static SectionType? ParseSection(Tokenizer reader, List<Token> tokens, out bool syntaxError)
    {
        bool hasDateParts = false;
        bool hasDurationParts = false;
        bool hasGeneralPart = false;
        bool hasTextPart = false;
        tokens.Clear();
        string source = reader.Source;

        syntaxError = false;
        while (ReadToken(reader, out var token))
        {
            if (token.IsCharacter(source, ';'))
                break;

            if (token.IsDatePart(source))
            {
                hasDateParts = true;
                hasDurationParts |= token.IsDurationPart(source);
                tokens.Add(token);
            }
            else if (token.IsGeneral(source))
            {
                hasGeneralPart = true;
                tokens.Add(token);
            }
            else if (token.IsCharacter(source, '@'))
            {
                hasTextPart = true;
                tokens.Add(token);
            }
            else if (source[token.Start] == '[')
            {
                ValidateCondition(reader, token.Start, token.Length);
            }
            else
            {
                tokens.Add(token);
            }
        }

        if (tokens.Count == 0)
            return null;

        if ((hasDateParts && (hasGeneralPart || hasTextPart)) ||
            (hasGeneralPart && (hasDateParts || hasTextPart)) ||
            (hasTextPart && (hasGeneralPart || hasDateParts)))
        {
            syntaxError = true;
            return null;
        }

        if (hasDateParts)
            return hasDurationParts ? SectionType.Duration : SectionType.Date;
        if (hasGeneralPart)
            return SectionType.General;
        if (hasTextPart)
            return SectionType.Text;
        if (IsFraction(reader, tokens))
            return SectionType.Fraction;

        int numberTokens = CountNumberTokens(source, tokens);
        if (numberTokens > 0 && numberTokens < tokens.Count && tokens[numberTokens].IsExponent(source))
            return SectionType.Exponential;
        if (numberTokens == tokens.Count)
            return SectionType.Number;

        syntaxError = true;
        return null;
    }

    private static int CountNumberTokens(string source, List<Token> tokens)
    {
        int index = 0;
        while (index < tokens.Count)
        {
            var token = tokens[index];
            if (!token.IsNumberLiteral(source) && source[token.Start] != '[')
                break;
            index++;
        }

        return index;
    }

    private static bool IsFraction(Tokenizer reader, List<Token> tokens)
    {
        string source = reader.Source;
        int index = 0;
        while (index < tokens.Count && !tokens[index].IsCharacter(source, '/'))
            index++;
        if (index == tokens.Count)
            return false;

        index++;
        while (index < tokens.Count)
        {
            var token = tokens[index];
            if (token.IsPlaceholder(source))
                return true;
            if (token.IsDigit19(source))
            {
                int first = index;
                while (index < tokens.Count && tokens[index].IsDigit09(source))
                    index++;

                // Constant denominators are validated even though their value is not needed.
                reader.ParseInt32(tokens, first, index - first);
                return true;
            }

            index++;
        }

        return false;
    }

    private static bool ReadToken(Tokenizer reader, out Token token)
    {
        int offset = reader.Position;
        if (ReadLiteral(reader) ||
            reader.ReadEnclosed('[', ']') ||
            reader.ReadOneOf("#?,!&%+-$\u20AC\u00A30123456789{}():;/.@ ") ||
            reader.ReadString("e+", true) ||
            reader.ReadString("e-", true) ||
            reader.ReadString("General", true) ||
            reader.ReadString("am/pm", true) ||
            reader.ReadString("a/p", true) ||
            reader.ReadOneOrMore('y') ||
            reader.ReadOneOrMore('Y') ||
            reader.ReadOneOrMore('m') ||
            reader.ReadOneOrMore('M') ||
            reader.ReadOneOrMore('d') ||
            reader.ReadOneOrMore('D') ||
            reader.ReadOneOrMore('h') ||
            reader.ReadOneOrMore('H') ||
            reader.ReadOneOrMore('s') ||
            reader.ReadOneOrMore('S') ||
            reader.ReadOneOrMore('g') ||
            reader.ReadOneOrMore('G'))
        {
            token = new Token(offset, reader.Position - offset);
            return true;
        }

        // Unrecognized characters remain implicit one-character tokens, not syntax errors.
        if (reader.Position < reader.Length)
        {
            reader.Advance();
            token = new Token(offset, reader.Position - offset);
            return true;
        }

        token = default;
        return false;
    }

    private static bool ReadLiteral(Tokenizer reader)
    {
        if (reader.Peek() is '\\' or '*' or '_')
        {
            reader.Advance(2);
            return true;
        }

        return reader.ReadEnclosed('"', '"');
    }

    private static void ValidateCondition(Tokenizer reader, int startIndex, int length)
    {
        // An unmatched '[' was previously passed to Substring with a negative length.
#if NET8_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfLessThan(length, 2);
#else
        if (length < 2)
            throw new ArgumentOutOfRangeException(nameof(length));
#endif

        string source = reader.Source;
        int position = startIndex + 1;
        int end = startIndex + length - 1;
        if (position == end || source[position] is not ('<' or '>' or '='))
            return;

        char first = source[position++];
        if (position < end && ((first is '<' or '>' && source[position] == '=') || (first == '<' && source[position] == '>')))
            position++;

        int start = position;
        if (position < end && source[position] == '-')
            position++;
        ReadDigits(source, ref position, end);
        if (position < end && source[position] == '.')
        {
            position++;
            ReadDigits(source, ref position, end);
        }

        if (position + 1 < end && char.ToLower(source[position], CultureInfo.InvariantCulture) == 'e' && source[position + 1] is '+' or '-')
        {
            position += 2;
            int exponentStart = position;
            ReadDigits(source, ref position, end);
            if (position == exponentStart)
                return;
        }

        reader.ParseDouble(start, position - start);
    }

    private static void ReadDigits(string source, ref int position, int end)
    {
        while (position < end && source[position] is >= '0' and <= '9')
            position++;
    }
}
