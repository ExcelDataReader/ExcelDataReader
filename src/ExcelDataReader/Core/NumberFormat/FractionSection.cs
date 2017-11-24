using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.NumberFormat
{
    internal class FractionSection
    {
        public List<string> IntegerPart { get; set; }

        public List<string> Numerator { get; set; }

        public List<string> DenominatorPrefix { get; set; }

        public List<string> Denominator { get; set; }

        public int DenominatorConstant { get; set; }

        public List<string> DenominatorSuffix { get; set; }

        public List<string> FractionSuffix { get; set; }

        public static bool TryParse(List<string> tokens, out FractionSection format)
        {
            List<string> numeratorParts = null;
            List<string> denominatorParts = null;

            for (var i = 0; i < tokens.Count; i++)
            {
                var part = tokens[i];
                if (part == "/")
                {
                    numeratorParts = tokens.GetRange(0, i);
                    i++;
                    denominatorParts = tokens.GetRange(i, tokens.Count - i);
                    break;
                }
            }

            if (numeratorParts == null)
            {
                format = null;
                return false;
            }

            GetNumerator(numeratorParts, out var integerPart, out var numeratorPart);

            if (!TryGetDenominator(denominatorParts, out var denominatorPrefix, out var denominatorPart, out var denominatorConstant, out var denominatorSuffix, out var fractionSuffix))
            {
                format = null;
                return false;
            }

            format = new FractionSection()
            {
                IntegerPart = integerPart,
                Numerator = numeratorPart,
                DenominatorPrefix = denominatorPrefix,
                Denominator = denominatorPart,
                DenominatorConstant = denominatorConstant,
                DenominatorSuffix = denominatorSuffix,
                FractionSuffix = fractionSuffix
            };

            return true;
        }

        private static void GetNumerator(List<string> tokens, out List<string> integerPart, out List<string> numeratorPart)
        {
            var hasPlaceholder = false;
            var hasSpace = false;
            var hasIntegerPart = false;
            var numeratorIndex = -1;
            var index = tokens.Count - 1;
            while (index >= 0)
            {
                var token = tokens[index];
                if (Token.IsPlaceholder(token))
                {
                    hasPlaceholder = true;

                    if (hasSpace)
                    {
                        hasIntegerPart = true;
                        break;
                    }
                }
                else
                {
                    if (hasPlaceholder && !hasSpace)
                    {
                        // First time we get here marks the end of the integer part
                        hasSpace = true;
                        numeratorIndex = index + 1;
                    }
                }

                index--;
            }

            if (hasIntegerPart)
            {
                integerPart = tokens.GetRange(0, numeratorIndex);
                numeratorPart = tokens.GetRange(numeratorIndex, tokens.Count - numeratorIndex);
            }
            else
            {
                integerPart = null;
                numeratorPart = tokens;
            }
        }

        private static bool TryGetDenominator(List<string> tokens, out List<string> denominatorPrefix, out List<string> denominatorPart, out int denominatorConstant, out List<string> denominatorSuffix, out List<string> fractionSuffix)
        {
            var index = 0;
            var hasPlaceholder = false;
            var hasConstant = false;

            var constant = new StringBuilder();

            // Read literals until the first number placeholder or digit
            while (index < tokens.Count)
            {
                var token = tokens[index];
                if (Token.IsPlaceholder(token))
                {
                    hasPlaceholder = true;
                    break;
                }
                else if (Token.IsDigit19(token))
                {
                    hasConstant = true;
                    break;
                }

                index++;
            }

            if (!hasPlaceholder && !hasConstant)
            {
                denominatorPrefix = null;
                denominatorPart = null;
                denominatorConstant = 0;
                denominatorSuffix = null;
                fractionSuffix = null;
                return false;
            }

            // The denominator starts here, keep the index
            var denominatorIndex = index;

            // Read placeholders or digits in sequence
            while (index < tokens.Count)
            {
                var token = tokens[index];
                if (hasPlaceholder && Token.IsPlaceholder(token))
                {
                    // OK
                }
                else
                if (hasConstant && Token.IsDigit09(token))
                {
                    constant.Append(token);
                }
                else
                {
                    break;
                }

                index++;
            }

            // 'index' is now at the first token after the denominator placeholders.
            // The remaining, if anything, is to be treated in one or two parts:
            // Any ultimately terminating literals are considered the "Fraction suffix".
            // Anything between the denominator and the fraction suffix is the "Denominator suffix".
            // Placeholders in the denominator suffix are treated as insignificant zeros.

            // Scan backwards to determine the fraction suffix
            int fractionSuffixIndex = tokens.Count;
            while (fractionSuffixIndex > index)
            {
                var token = tokens[fractionSuffixIndex - 1];
                if (Token.IsPlaceholder(token))
                {
                    break;
                }

                fractionSuffixIndex--;
            }

            // Finally extract the detected token ranges
            if (denominatorIndex > 0)
                denominatorPrefix = tokens.GetRange(0, denominatorIndex);
            else
                denominatorPrefix = null;

            if (hasConstant)
                denominatorConstant = int.Parse(constant.ToString());
            else
                denominatorConstant = 0;

            denominatorPart = tokens.GetRange(denominatorIndex, index - denominatorIndex);

            if (index < fractionSuffixIndex)
                denominatorSuffix = tokens.GetRange(index, fractionSuffixIndex - index);
            else
                denominatorSuffix = null;

            if (fractionSuffixIndex < tokens.Count)
                fractionSuffix = tokens.GetRange(fractionSuffixIndex, tokens.Count - fractionSuffixIndex);
            else
                fractionSuffix = null;

            return true;
        }
    }
}