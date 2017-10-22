using System.Collections.Generic;

namespace ExcelDataReader.Core.NumberFormat
{
    internal class ExponentialSection
    {
        public List<string> BeforeDecimal { get; set; }

        public bool DecimalSeparator { get; set; }

        public List<string> AfterDecimal { get; set; }

        public string ExponentialToken { get; set; }

        public List<string> Power { get; set; }

        public static bool TryParse(List<string> tokens, out ExponentialSection format)
        {
            format = null;

            string exponentialToken;

            int partCount = Parser.ParseNumberTokens(tokens, 0, out var beforeDecimal, out var decimalSeparator, out var afterDecimal);

            if (partCount == 0)
                return false;

            int position = partCount;
            if (position < tokens.Count && Token.IsExponent(tokens[position]))
            {
                exponentialToken = tokens[position];
                position++;
            }
            else
            {
                return false;
            }

            format = new ExponentialSection()
            {
                BeforeDecimal = beforeDecimal,
                DecimalSeparator = decimalSeparator,
                AfterDecimal = afterDecimal,
                ExponentialToken = exponentialToken,
                Power = tokens.GetRange(position, tokens.Count - position)
            };

            return true;
        }
    }
}