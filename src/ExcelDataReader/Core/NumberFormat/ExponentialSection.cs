namespace ExcelDataReader.Core.NumberFormat;

internal sealed class ExponentialSection
{
    public required List<string> BeforeDecimal { get; init; }

    public required bool DecimalSeparator { get; init; }

    public required List<string> AfterDecimal { get; init; }

    public required string ExponentialToken { get; init; }

    public required List<string> Power { get; init; }

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