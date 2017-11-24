using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelDataReader.Core.NumberFormat
{
    internal static class Parser
    {
        public static Section ParseSection(Tokenizer reader, out bool syntaxError)
        {
            bool hasDateParts = false;
            bool hasDurationParts = false;
            bool hasGeneralPart = false;
            bool hasTextPart = false;
            Condition condition = null;
            Color color = null;
            string token;
            List<string> tokens = new List<string>();

            syntaxError = false;
            while ((token = ReadToken(reader, out syntaxError)) != null)
            {
                if (token == ";")
                    break;

                if (Token.IsDatePart(token))
                {
                    hasDateParts |= true;
                    hasDurationParts |= Token.IsDurationPart(token);
                    tokens.Add(token);
                }
                else if (Token.IsGeneral(token))
                {
                    hasGeneralPart |= true;
                    tokens.Add(token);
                }
                else if (token == "@")
                {
                    hasTextPart |= true;
                    tokens.Add(token);
                }
                else if (token.StartsWith("["))
                {
                    // Does not add to tokens. Absolute/elapsed time tokens
                    // also start with '[', but handled as date part above
                    var expression = token.Substring(1, token.Length - 2);
                    if (TryParseCondition(expression, out var parseCondition))
                        condition = parseCondition;
                    else if (TryParseColor(expression, out var parseColor))
                        color = parseColor;
                }
                else
                {
                    tokens.Add(token);
                }
            }

            if (syntaxError || tokens.Count == 0)
            {
                return null;
            }

            if (
                (hasDateParts && (hasGeneralPart || hasTextPart)) ||
                (hasGeneralPart && (hasDateParts || hasTextPart)) ||
                (hasTextPart && (hasGeneralPart || hasDateParts)))
            {
                // Cannot mix date, general and/or text parts
                syntaxError = true;
                return null;
            }

            SectionType type;
            FractionSection fraction = null;
            ExponentialSection exponential = null;
            DecimalSection number = null;
            List<string> generalTextDateDuration = null;

            if (hasDateParts)
            {
                if (hasDurationParts)
                {
                    type = SectionType.Duration;
                    generalTextDateDuration = tokens;
                }
                else
                {
                    type = SectionType.Date;
                    ParseDate(tokens, out generalTextDateDuration);
                }
            }
            else if (hasGeneralPart)
            {
                type = SectionType.General;
                generalTextDateDuration = tokens;
            }
            else if (hasTextPart)
            {
                type = SectionType.Text;
                generalTextDateDuration = tokens;
            }
            else if (FractionSection.TryParse(tokens, out fraction))
            {
                type = SectionType.Fraction;
            }
            else if (ExponentialSection.TryParse(tokens, out exponential))
            {
                type = SectionType.Exponential;
            }
            else if (DecimalSection.TryParse(tokens, out number))
            {
                type = SectionType.Number;
            }
            else
            {
                // Unable to parse format string
                syntaxError = true;
                return null;
            }

            return new Section()
            {
                Type = type,
                Color = color,
                Condition = condition,
                Fraction = fraction,
                Exponential = exponential,
                Number = number,
                GeneralTextDateDurationParts = generalTextDateDuration
            };
        }

        /// <summary>
        /// Parses as many placeholders and literals needed to format a number with optional decimals. 
        /// Returns number of tokens parsed, or 0 if the tokens didn't form a number.
        /// </summary>
        internal static int ParseNumberTokens(List<string> tokens, int startPosition, out List<string> beforeDecimal, out bool decimalSeparator, out List<string> afterDecimal)
        {
            beforeDecimal = null;
            afterDecimal = null;
            decimalSeparator = false;

            List<string> remainder = new List<string>();
            var index = 0;
            for (index = 0; index < tokens.Count; ++index)
            {
                var token = tokens[index];
                if (token == "." && beforeDecimal == null)
                {
                    decimalSeparator = true;
                    beforeDecimal = tokens.GetRange(0, index); // TODO: why not remainder? has only valid tokens...

                    remainder = new List<string>();
                }
                else if (Token.IsNumberLiteral(token))
                {
                    remainder.Add(token);
                }
                else if (token.StartsWith("["))
                {
                    // ignore
                }
                else
                {
                    break;
                }
            }

            if (remainder.Count > 0)
            {
                if (beforeDecimal != null)
                {
                    afterDecimal = remainder;
                }
                else
                {
                    beforeDecimal = remainder;
                }
            }
            
            return index;
        }

        private static void ParseDate(List<string> tokens, out List<string> result)
        {
            // if tokens form .0 through .000.., combine to single subsecond token
            result = new List<string>();
            for (var i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];
                if (token == ".")
                {
                    var zeros = 0;
                    while (i + 1 < tokens.Count && tokens[i + 1] == "0")
                    {
                        i++;
                        zeros++;
                    }

                    if (zeros > 0)
                        result.Add("." + new string('0', zeros));
                    else
                        result.Add(".");
                }
                else
                {
                    result.Add(token);
                }
            }
        }

        private static string ReadToken(Tokenizer reader, out bool syntaxError)
        {
            var offset = reader.Position;
            if (
                ReadLiteral(reader) ||
                reader.ReadEnclosed('[', ']') ||

                // Symbols
                reader.ReadOneOf("#?,!&%+-$€£0123456789{}():;/.@ ") ||
                reader.ReadString("e+", true) ||
                reader.ReadString("e-", true) ||
                reader.ReadString("General", true) ||

                // Date
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
                syntaxError = false;
                var length = reader.Position - offset;
                return reader.Substring(offset, length);
            }

            syntaxError = reader.Position < reader.Length;
            return null;
        }

        private static bool ReadLiteral(Tokenizer reader)
        {
            if (reader.Peek() == '\\' || reader.Peek() == '*' || reader.Peek() == '_')
            {
                reader.Advance(2);
                return true;
            }
            else if (reader.ReadEnclosed('"', '"'))
            {
                return true;
            }

            return false;
        }

        private static bool TryParseCondition(string token, out Condition result)
        {
            var tokenizer = new Tokenizer(token);

            if (tokenizer.ReadString("<=") ||
                tokenizer.ReadString("<>") ||
                tokenizer.ReadString("<") ||
                tokenizer.ReadString(">=") ||
                tokenizer.ReadString(">") ||
                tokenizer.ReadString("="))
            {
                var conditionPosition = tokenizer.Position;
                var op = tokenizer.Substring(0, conditionPosition);

                if (ReadConditionValue(tokenizer))
                {
                    var valueString = tokenizer.Substring(conditionPosition, tokenizer.Position - conditionPosition);

                    result = new Condition()
                    {
                        Operator = op,
                        Value = double.Parse(valueString, CultureInfo.InvariantCulture)
                    };
                    return true;
                }
            }

            result = null;
            return false;
        }

        private static bool ReadConditionValue(Tokenizer tokenizer)
        {
            // NFPartCondNum = [ASCII-HYPHEN-MINUS] NFPartIntNum [INTL-CHAR-DECIMAL-SEP NFPartIntNum] [NFPartExponential NFPartIntNum]
            tokenizer.ReadString("-");
            while (tokenizer.ReadOneOf("0123456789"))
            {
            }

            if (tokenizer.ReadString("."))
            {
                while (tokenizer.ReadOneOf("0123456789"))
                {
                }
            }

            if (tokenizer.ReadString("e+", true) || tokenizer.ReadString("e-", true))
            {
                if (tokenizer.ReadOneOf("0123456789"))
                {
                    while (tokenizer.ReadOneOf("0123456789"))
                    {
                    }
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        private static bool TryParseColor(string token, out Color color)
        {
            // TODO: Color1..59
            var tokenizer = new Tokenizer(token);
            if (
                tokenizer.ReadString("black", true) ||
                tokenizer.ReadString("blue", true) ||
                tokenizer.ReadString("cyan", true) ||
                tokenizer.ReadString("green", true) ||
                tokenizer.ReadString("magenta", true) ||
                tokenizer.ReadString("red", true) ||
                tokenizer.ReadString("white", true) ||
                tokenizer.ReadString("yellow", true))
            {
                color = new Color()
                {
                    Value = tokenizer.Substring(0, tokenizer.Position)
                };
                return true;
            }

            color = null;
            return false;
        }
    }
}
