namespace ExcelDataReader.Core.NumberFormat;

internal readonly record struct Token(int Start, int Length)
{
    public bool IsCharacter(string source, char value) => Length == 1 && source[Start] == value;

    public bool IsExponent(string source) => EqualsIgnoreCase(source, "e+") || EqualsIgnoreCase(source, "e-");

    public bool IsNumberLiteral(string source)
    {
        char first = source[Start];
        if (first is '_' or '\\' or '"' or '*')
            return true;

        return Length == 1 && first is '0' or '#' or '?' or '.' or ',' or '!' or '&' or '%' or '+' or '-' or '$' or '\u20AC' or '\u00A3' or
            '1' or '2' or '3' or '4' or '5' or '6' or '7' or '8' or '9' or '{' or '}' or '(' or ')' or ' ';
    }

    public bool IsPlaceholder(string source) => Length == 1 && source[Start] is '0' or '#' or '?';

    public bool IsGeneral(string source) => EqualsIgnoreCase(source, "general");

    public bool IsDatePart(string source)
    {
        return StartsWithIgnoreCase(source, "y") ||
            StartsWithIgnoreCase(source, "m") ||
            StartsWithIgnoreCase(source, "d") ||
            StartsWithIgnoreCase(source, "s") ||
            StartsWithIgnoreCase(source, "h") ||
            (StartsWithIgnoreCase(source, "g") && !IsGeneral(source)) ||
            EqualsIgnoreCase(source, "am/pm") ||
            EqualsIgnoreCase(source, "a/p") ||
            IsDurationPart(source);
    }

    public bool IsDurationPart(string source) =>
        StartsWithIgnoreCase(source, "[h") || StartsWithIgnoreCase(source, "[m") || StartsWithIgnoreCase(source, "[s");

    public bool IsDigit09(string source) => Length == 1 && source[Start] is >= '0' and <= '9';

    public bool IsDigit19(string source) => Length == 1 && source[Start] is >= '1' and <= '9';

    private bool EqualsIgnoreCase(string source, string value) => Length == value.Length && StartsWithIgnoreCase(source, value);

    private bool StartsWithIgnoreCase(string source, string value)
    {
        if (Length < value.Length)
            return false;
        for (int i = 0; i < value.Length; i++)
        {
            if (!Helpers.EqualsAsciiIgnoreCase(source[Start + i], value[i]))
                return false;
        }

        return true;
    }
}