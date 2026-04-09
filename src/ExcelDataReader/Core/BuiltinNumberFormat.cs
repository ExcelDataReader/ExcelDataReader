#nullable enable

using System.Globalization;
using System.Text;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core;

internal static class BuiltinNumberFormat
{
    private static Dictionary<int, NumberFormatString> Formats { get; } = new()
    {
        { 0, new NumberFormatString("General") },
        { 1, new NumberFormatString("0") },
        { 2, new NumberFormatString("0.00") },
        { 3, new NumberFormatString("#,##0") },
        { 4, new NumberFormatString("#,##0.00") },
        { 5, new NumberFormatString("\"$\"#,##0_);(\"$\"#,##0)") },
        { 6, new NumberFormatString("\"$\"#,##0_);[Red](\"$\"#,##0)") },
        { 7, new NumberFormatString("\"$\"#,##0.00_);(\"$\"#,##0.00)") },
        { 8, new NumberFormatString("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)") },
        { 9, new NumberFormatString("0%") },
        { 10, new NumberFormatString("0.00%") },
        { 11, new NumberFormatString("0.00E+00") },
        { 12, new NumberFormatString("# ?/?") },
        { 13, new NumberFormatString("# ??/??") },
        { 14, new NumberFormatString("d/m/yyyy") },
        { 15, new NumberFormatString("d-mmm-yy") },
        { 16, new NumberFormatString("d-mmm") },
        { 17, new NumberFormatString("mmm-yy") },
        { 18, new NumberFormatString("h:mm AM/PM") },
        { 19, new NumberFormatString("h:mm:ss AM/PM") },
        { 20, new NumberFormatString("h:mm") },
        { 21, new NumberFormatString("h:mm:ss") },
        { 22, new NumberFormatString("m/d/yy h:mm") },

        // 23..36 international/unused
        { 37, new NumberFormatString("#,##0_);(#,##0)") },
        { 38, new NumberFormatString("#,##0_);[Red](#,##0)") },
        { 39, new NumberFormatString("#,##0.00_);(#,##0.00)") },
        { 40, new NumberFormatString("#,##0.00_);[Red](#,##0.00)") },
        { 41, new NumberFormatString("_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)") },
        { 42, new NumberFormatString("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)") },
        { 43, new NumberFormatString("_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)") },
        { 44, new NumberFormatString("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)") },
        { 45, new NumberFormatString("mm:ss") },
        { 46, new NumberFormatString("[h]:mm:ss") },
        { 47, new NumberFormatString("mm:ss.0") },
        { 48, new NumberFormatString("##0.0E+0") },
        { 49, new NumberFormatString("@") },

        // Locale-specific built-in formats.
        // Indices 27–36 and 50–58 vary by locale (ja-jp, ko-kr, zh-tw, zh-cn) but are
        // always date/time formats. The ja-jp strings are stored as a representative
        // fallback; files from other locales may display differently via GetNumberFormatString()
        // but will still be correctly identified as date/time values.
        { 27, new NumberFormatString(@"[$-411]ge.m.d") },
        { 28, new NumberFormatString("[$-411]ggge\"年\"m\"月\"d\"日\"") },
        { 29, new NumberFormatString("[$-411]ggge\"年\"m\"月\"d\"日\"") },
        { 30, new NumberFormatString("m/d/yy") },
        { 31, new NumberFormatString("yyyy\"年\"m\"月\"d\"日\"") },
        { 32, new NumberFormatString("h\"時\"mm\"分\"") },
        { 33, new NumberFormatString("h\"時\"mm\"分\"ss\"秒\"") },
        { 34, new NumberFormatString("yyyy\"年\"m\"月\"") },
        { 35, new NumberFormatString("m\"月\"d\"日\"") },
        { 36, new NumberFormatString(@"[$-411]ge.m.d") },
        { 50, new NumberFormatString(@"[$-411]ge.m.d") },
        { 51, new NumberFormatString("[$-411]ggge\"年\"m\"月\"d\"日\"") },
        { 52, new NumberFormatString("yyyy\"年\"m\"月\"") },
        { 53, new NumberFormatString("m\"月\"d\"日\"") },
        { 54, new NumberFormatString("[$-411]ggge\"年\"m\"月\"d\"日\"") },
        { 55, new NumberFormatString("yyyy\"年\"m\"月\"") },
        { 56, new NumberFormatString("m\"月\"d\"日\"") },
        { 57, new NumberFormatString(@"[$-411]ge.m.d") },
        { 58, new NumberFormatString("[$-411]ggge\"年\"m\"月\"d\"日\"") },

        // 59–62 and 67–70 are Thai-digit number formats; 71–81 are Thai date/time formats.
        // ASCII equivalents are used so the format strings parse correctly.
        { 59, new NumberFormatString("0") },
        { 60, new NumberFormatString("0.00") },
        { 61, new NumberFormatString("#,##0") },
        { 62, new NumberFormatString("#,##0.00") },
        { 67, new NumberFormatString("0%") },
        { 68, new NumberFormatString("0.00%") },
        { 69, new NumberFormatString("# ?/?") },
        { 70, new NumberFormatString("# ??/??") },
        { 71, new NumberFormatString("d/m/yyyy") },
        { 72, new NumberFormatString("d-mmm-yy") },
        { 73, new NumberFormatString("d-mmm") },
        { 74, new NumberFormatString("mmm-yy") },
        { 75, new NumberFormatString("h:mm") },
        { 76, new NumberFormatString("h:mm:ss") },
        { 77, new NumberFormatString("d/m/yyyy h:mm") },
        { 78, new NumberFormatString("mm:ss") },
        { 79, new NumberFormatString("[h]:mm:ss") },
        { 80, new NumberFormatString("mm:ss.0") },
        { 81, new NumberFormatString("d/m/yyyy") },
    };

    public static NumberFormatString? GetBuiltinNumberFormat(int numFmtId)
    {
        if (Formats.TryGetValue(numFmtId, out var result))
            return result;

        return null;
    }

    public static NumberFormatString? GetBuiltinNumberFormat(int numFmtId, CultureInfo culture)
    {
        if (numFmtId == 14)
            return new NumberFormatString(DatePatternToExcel(culture.DateTimeFormat.ShortDatePattern));

        if (numFmtId == 15 || numFmtId == 16 || numFmtId == 17)
        {
            var sep = EscapeExcelLiteral(culture.DateTimeFormat.DateSeparator);
            return numFmtId switch
            {
                15 => new NumberFormatString($"d{sep}mmm{sep}yy"),
                16 => new NumberFormatString($"d{sep}mmm"),
                _ => new NumberFormatString($"mmm{sep}yy"),
            };
        }

        if (numFmtId == 22)
        {
            var date = DatePatternToExcel(culture.DateTimeFormat.ShortDatePattern);
            var time = TimePatternToExcel(culture.DateTimeFormat.ShortTimePattern);
            return new NumberFormatString($"{date} {time}");
        }

        if (Formats.TryGetValue(numFmtId, out var result))
            return result;

        return null;
    }

    // Wraps a date separator in double quotes if it contains any letter that would be
    // misinterpreted as a date/time format specifier in an Excel format string.
    // In practice, date separators are always punctuation (/, ., -) so quoting is rarely needed.
    private static string EscapeExcelLiteral(string s)
    {
        foreach (char c in s)
        {
            if (char.IsLetter(c) || c == '"')
                return $"\"{s}\"";
        }

        return s;
    }

    // Converts a .NET date pattern (e.g. "M/d/yyyy") to an Excel format string (e.g. "m/d/yyyy").
    // .NET uses uppercase M for months; Excel uses lowercase m.
    // Single-quoted literals 'text' are converted to double-quoted "text" without modifying their content.
    // Backslash-escaped characters \x are converted to "x".
    private static string DatePatternToExcel(string dotNetPattern)
    {
        var sb = new StringBuilder(dotNetPattern.Length + 4);
        for (int i = 0; i < dotNetPattern.Length; i++)
        {
            char c = dotNetPattern[i];
            if (c == '\'')
            {
                sb.Append('"');
                i++;
                while (i < dotNetPattern.Length && dotNetPattern[i] != '\'')
                    sb.Append(dotNetPattern[i++]);
                sb.Append('"');
            }
            else if (c == '\\' && i + 1 < dotNetPattern.Length)
            {
                sb.Append('"').Append(dotNetPattern[++i]).Append('"');
            }
            else if (c == 'M')
            {
                sb.Append('m');
            }
            else
            {
                sb.Append(c);
            }
        }

        return sb.ToString();
    }

    // Converts a .NET time pattern (e.g. "h:mm tt") to an Excel format string (e.g. "h:mm AM/PM").
    // .NET uses uppercase H for 24-hour clock; Excel uses lowercase h.
    // .NET uses tt/t for AM/PM designator; Excel uses AM/PM and A/P.
    // Single-quoted literals and backslash escapes are handled as in DatePatternToExcel.
    private static string TimePatternToExcel(string dotNetPattern)
    {
        var sb = new StringBuilder(dotNetPattern.Length + 8);
        for (int i = 0; i < dotNetPattern.Length; i++)
        {
            char c = dotNetPattern[i];
            if (c == '\'')
            {
                sb.Append('"');
                i++;
                while (i < dotNetPattern.Length && dotNetPattern[i] != '\'')
                    sb.Append(dotNetPattern[i++]);
                sb.Append('"');
            }
            else if (c == '\\' && i + 1 < dotNetPattern.Length)
            {
                sb.Append('"').Append(dotNetPattern[++i]).Append('"');
            }
            else if (c == 'H')
            {
                sb.Append('h');
            }
            else if (c == 't' && i + 1 < dotNetPattern.Length && dotNetPattern[i + 1] == 't')
            {
                sb.Append("AM/PM");
                i++;
            }
            else if (c == 't')
            {
                sb.Append("A/P");
            }
            else
            {
                sb.Append(c);
            }
        }

        return sb.ToString();
    }
}
