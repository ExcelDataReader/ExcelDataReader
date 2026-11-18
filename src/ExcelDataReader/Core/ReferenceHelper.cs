namespace ExcelDataReader.Core;

internal static class ReferenceHelper
{
    /// <summary>
    /// Logic for the Excel dimensions. Ex: A15.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <param name="column">The column, 1-based.</param>
    /// <param name="row">The row, 1-based.</param>
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
    public static bool ParseReference(ReadOnlySpan<char> value, out int column, out int row)
#else
    public static bool ParseReference(string value, out int column, out int row)
#endif
    {
        column = 0;
        row = 0;
#if !(NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER)
        if (value == null)
            return false;
#endif
        if (!TryParseColumn(value, out column, out int position))
            return false;

        return TryParseDecInt(value, position, out row);
    }

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
    private static bool TryParseColumn(ReadOnlySpan<char> value, out int column, out int position)
#else
    private static bool TryParseColumn(string value, out int column, out int position)
#endif
    {
        column = 0;
        position = 0;
        while (position < value.Length)
        {
            var c = value[position];
            int letter = (c | 0x20) - 'a';
            if (char.IsAsciiLetter(c))
            {
                if (column > (int.MaxValue - letter - 1) / 26)
                {
                    column = 0;
                    return false;
                }

                position++;
                column *= 26;
                column += letter + 1;
                continue;
            }

            if (char.IsAsciiDigit(c))
                return position > 0;

            column = 0;
            return false;
        }

        return position > 0;
    }

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
    private static bool TryParseDecInt(ReadOnlySpan<char> s, int startIndex, out int result)
#else
    private static bool TryParseDecInt(string s, int startIndex, out int result)
#endif
    {
        if (startIndex >= s.Length)
            goto Fail;

        long r = 0;
        for (int i = startIndex; i < s.Length; i++)
        {
            char num = s[i];
            if (!char.IsAsciiDigit(num))
                goto Fail;
            r = r * 10 + (num - '0');
            if (r > int.MaxValue)
                goto Fail;
        }

        result = (int)r;
        return true;

    Fail:
        result = 0;
        return false;
    }
}