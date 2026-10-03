using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader.Misc;

namespace ExcelDataReader.Core;

/// <summary>
/// Helpers class.
/// </summary>
internal static class Helpers
{
    #if !(NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER)
    private static readonly Regex EscapeRegexInstance = new("_x([0-9A-F]{4,4})_", RegexOptions.Compiled);
    #endif

    private static readonly char[] SingleByteEncodingHelper = ['a'];

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

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        return DecodeEscapeChars(input, index);
#else
        return EscapeRegex().Replace(input, m => ((char)uint.Parse(m.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString());
#endif
    }

    public static object ConvertFromOATime(double value, bool date1904)
    {
        var dateValue = AdjustOADateTime(value, date1904);
        if (IsValidOADateTime(dateValue))
            return DateTimeHelper.FromOADate(dateValue);
        return value;
    }

    public static object ConvertFromOATime(int value, bool date1904)
    {
        var dateValue = AdjustOADateTime(value, date1904);
        if (IsValidOADateTime(dateValue))
            return DateTimeHelper.FromOADate(dateValue);
        return value;
    }

    public static bool StringStartsWith(string value, char start)
    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        return value.StartsWith(start);
#else
        return value.Length > 0 && value[0] == start;
#endif
    }

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
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

        return string.Create(input.Length - matches * 6, (Input: input, FirstMatch: firstMatch), static (destination, state) =>
        {
            int start = 0;
            int written = 0;
            int index = state.FirstMatch;
            while (index >= 0)
            {
                bool valid = TryReadEscape(state.Input, index, out char value);
                if (valid)
                {
                    state.Input.AsSpan(start, index - start).CopyTo(destination[written..]);
                    written += index - start;
                    destination[written++] = value;
                    start = index + 7;
                }

                index = state.Input.IndexOf("_x", index + (valid ? 7 : 2), StringComparison.Ordinal);
            }

            state.Input.AsSpan(start).CopyTo(destination[written..]);
        });
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
#endif
    
    /// <summary>
    /// Convert a double from Excel to an OA DateTime double. 
    /// The returned value is normalized to the '1900' date mode and adjusted for the 1900 leap year bug.
    /// </summary>
    private static double AdjustOADateTime(double value, bool date1904)
    {
        if (!date1904)
        {
            // Workaround for 1900 leap year bug in Excel
            if (value is >= 0.0 and < 60.0)
            {
                return value + 1;
            }
        }
        else
        {
            return value + 1462.0;
        }

        return value;
    }

    private static bool IsValidOADateTime(double value)
    {
        return value is > DateTimeHelper.OADateMinAsDouble and < DateTimeHelper.OADateMaxAsDouble;
    }
    
#if !(NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER)
    private static Regex EscapeRegex() => EscapeRegexInstance;
#endif
}