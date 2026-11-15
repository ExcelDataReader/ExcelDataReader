using System.Text;

namespace ExcelDataReader.Core;

/// <summary>
/// Helpers class.
/// </summary>
internal static class Helpers
{
    private static readonly char[] SingleByteEncodingHelper = ['a'];
    private static readonly object True = true;
    private static readonly object False = false;

    public static object BoxBoolean(bool value) => value ? True : False;

    public static bool EqualsAsciiIgnoreCase(char left, char right)
    {
        int folded = left | 0x20;
        return left == right || ((uint)(folded - 'a') <= 'z' - 'a' && folded == (right | 0x20));
    }

    /// <summary>
    /// Determines whether the encoding is single byte or not.
    /// </summary>
    /// <param name="encoding">The encoding.</param>
    /// <returns>
    ///     <see langword="true"/> if the specified encoding is single byte; otherwise, <see langword="false"/>.
    /// </returns>
    public static bool IsSingleByteEncoding(Encoding encoding)
    {
        return encoding.GetByteCount(SingleByteEncodingHelper) == 1;
    }

    public static string ConvertEscapeChars(string input)
    {
        // Fast rejection: avoid scanning escape candidates for typical spreadsheet strings.
        int index = input.IndexOf("_x", StringComparison.Ordinal);
        if (index < 0)
            return input;

        return DecodeEscapeChars(input, index);
    }

    public static object ConvertFromOATime(double value, bool date1904)
    {
        int dayOffset = GetOADateOffset(value, date1904);
        if (IsValidOADateTime(value + dayOffset))
            return DateTimeHelper.FromOADate(value, dayOffset);
        return value;
    }

    public static object ConvertFromOATime(int value, bool date1904)
    {
        int dayOffset = GetOADateOffset(value, date1904);
        if (IsValidOADateTime((double)value + dayOffset))
            return DateTimeHelper.FromOADate(value, dayOffset);
        return value;
    }

    private static string DecodeEscapeChars(string input, int index)
    {
        int firstMatch = -1;
        int matches = 0;
        while (index >= 0)
        {
            bool valid = TryReadEscape(input, index, out _);
            if (valid)
            {
                if (firstMatch < 0)
                    firstMatch = index;
                matches++;
            }

            index = input.IndexOf("_x", index + (valid ? 7 : 2), StringComparison.Ordinal);
        }

        if (matches == 0)
            return input;

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        return string.Create(
            input.Length - matches * 6,
            (Input: input, FirstMatch: firstMatch),
            static (destination, state) => WriteEscapeChars(state.Input, state.FirstMatch, destination));
#else
        var destination = new char[input.Length - matches * 6];
        WriteEscapeChars(input, firstMatch, destination);
        return new string(destination);
#endif
    }

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
    private static void WriteEscapeChars(string input, int index, Span<char> destination)
#else
    private static void WriteEscapeChars(string input, int index, char[] destination)
#endif
    {
        int start = 0;
        int written = 0;
        while (index >= 0)
        {
            bool valid = TryReadEscape(input, index, out char value);
            if (valid)
            {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
                input.AsSpan(start, index - start).CopyTo(destination[written..]);
#else
                input.CopyTo(start, destination, written, index - start);
#endif
                written += index - start;
                destination[written++] = value;
                start = index + 7;
            }

            index = input.IndexOf("_x", index + (valid ? 7 : 2), StringComparison.Ordinal);
        }

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        input.AsSpan(start).CopyTo(destination[written..]);
#else
        input.CopyTo(start, destination, written, input.Length - start);
#endif
    }

    private static bool TryReadEscape(string input, int index, out char value)
    {
        value = default;
        if (index > input.Length - 7 || input[index + 6] != '_')
            return false;

        int code = 0;
        for (int i = index + 2; i < index + 6; i++)
        {
            int digit = input[i] switch
            {
                >= '0' and <= '9' => input[i] - '0',
                >= 'A' and <= 'F' => input[i] - 'A' + 10,
                _ => -1,
            };
            if (digit < 0)
                return false;
            code = (code << 4) | digit;
        }

        value = (char)code;
        return true;
    }

    /// <summary>
    /// Get the whole-day offset from an Excel date to its OA date.
    /// </summary>
    private static int GetOADateOffset(double value, bool date1904)
    {
        if (date1904)
            return 1462;
        return value is >= 0.0 and < 60.0 ? 1 : 0;
    }

    private static bool IsValidOADateTime(double value)
    {
        return value is > DateTimeHelper.OADateMinAsDouble and < DateTimeHelper.OADateMaxAsDouble;
    }
}