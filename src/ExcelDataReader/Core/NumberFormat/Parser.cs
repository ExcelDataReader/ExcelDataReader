namespace ExcelDataReader.Core.NumberFormat;

internal static class Parser
{
    public static SectionType? ParseSection(
        Tokenizer reader,
        List<Token> tokens,
        out bool syntaxError,
        out bool moreSections,
        out NumberFormatCondition? condition)
    {
        bool hasDateParts = false;
        bool hasDurationParts = false;
        bool hasGeneralPart = false;
        bool hasTextPart = false;
        tokens.Clear();
        string source = reader.Source;

        syntaxError = false;
        moreSections = false;
        condition = null;
        while (ReadToken(reader, out var token))
        {
            if (token.IsCharacter(source, ';'))
            {
                moreSections = true;
                break;
            }

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
                var parsedCondition = ValidateCondition(reader, token.Start, token.Length);
                condition ??= parsedCondition;
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

        // Dispatch on the first character: each alternative can only match when it starts with that character.
        int c = reader.Peek();
        bool matched = c switch
        {
            '\\' or '*' or '_' or '"' => ReadLiteral(reader),
            '[' => reader.ReadEnclosed('[', ']'),
            'e' or 'E' => reader.ReadString("e+", true) || reader.ReadString("e-", true),
            'g' or 'G' => reader.ReadString("General", true) || reader.ReadOneOrMore(c),
            'a' or 'A' => reader.ReadString("am/pm", true) || reader.ReadString("a/p", true),
            'y' or 'Y' or 'm' or 'M' or 'd' or 'D' or 'h' or 'H' or 's' or 'S' => reader.ReadOneOrMore(c),
            _ => reader.ReadOneOf("#?,!&%+-$\u20AC\u00A30123456789{}():;/.@ "),
        };

        if (matched)
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

    private static NumberFormatCondition? ValidateCondition(Tokenizer reader, int startIndex, int length)
    {
        // An unmatched '[' was previously passed to Substring with a negative length.
        ArgumentOutOfRangeException.ThrowIfLessThan(length, 2);

        string source = reader.Source;
        int position = startIndex + 1;
        int end = startIndex + length - 1;
        if (position == end || source[position] is not ('<' or '>' or '='))
            return null;

        char first = source[position++];
        ConditionOperator op = first switch
        {
            '<' => ConditionOperator.LessThan,
            '>' => ConditionOperator.GreaterThan,
            '=' => ConditionOperator.Equal,
            _ => throw new InvalidOperationException(),
        };
        if (position < end && ((first is '<' or '>' && source[position] == '=') || (first == '<' && source[position] == '>')))
        {
            op = (first, source[position]) switch
            {
                ('<', '=') => ConditionOperator.LessThanOrEqual,
                ('>', '=') => ConditionOperator.GreaterThanOrEqual,
                ('<', '>') => ConditionOperator.NotEqual,
                _ => op,
            };
            position++;
        }

        int start = position;
        if (position < end && source[position] == '-')
            position++;
        ReadDigits(source, ref position, end);
        if (position < end && source[position] == '.')
        {
            position++;
            ReadDigits(source, ref position, end);
        }

        if (position + 1 < end && source[position] is 'e' or 'E' && source[position + 1] is '+' or '-')
        {
            position += 2;
            int exponentStart = position;
            ReadDigits(source, ref position, end);
            if (position == exponentStart)
                return null;
        }

        double value = start == end && op == ConditionOperator.GreaterThanOrEqual
            ? 0
            : reader.ParseDouble(start, position - start);
        return new NumberFormatCondition(op, value);
    }

    private static void ReadDigits(string source, ref int position, int end)
    {
        while (position < end && char.IsAsciiDigit(source[position]))
            position++;
    }
}