#nullable enable

using System.Globalization;
using System.Text;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core;

internal static class BuiltinNumberFormat
{
    private static Dictionary<int, NumberFormatString> Formats { get; } = new()
    {
        { 0, Number("General") },
        { 1, Number("0") },
        { 2, Number("0.00") },
        { 3, Number("#,##0") },
        { 4, Number("#,##0.00") },
        { 5, Number("\"$\"#,##0_);(\"$\"#,##0)") },
        { 6, Number("\"$\"#,##0_);[Red](\"$\"#,##0)") },
        { 7, Number("\"$\"#,##0.00_);(\"$\"#,##0.00)") },
        { 8, Number("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)") },
        { 9, Number("0%") },
        { 10, Number("0.00%") },
        { 11, Number("0.00E+00") },
        { 12, Number("# ?/?") },
        { 13, Number("# ??/??") },
        { 14, Date("d/m/yyyy") },
        { 15, Date("d-mmm-yy") },
        { 16, Date("d-mmm") },
        { 17, Date("mmm-yy") },
        { 18, Date("h:mm AM/PM") },
        { 19, Date("h:mm:ss AM/PM") },
        { 20, Date("h:mm") },
        { 21, Date("h:mm:ss") },
        { 22, Date("m/d/yy h:mm") },

        // 23..26 international/unused
        { 37, Number("#,##0_);(#,##0)") },
        { 38, Number("#,##0_);[Red](#,##0)") },
        { 39, Number("#,##0.00_);(#,##0.00)") },
        { 40, Number("#,##0.00_);[Red](#,##0.00)") },
        { 41, Number("_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)") },
        { 42, Number("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)") },
        { 43, Number("_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)") },
        { 44, Number("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)") },
        { 45, Date("mm:ss") },
        { 46, Duration("[h]:mm:ss") },
        { 47, Date("mm:ss.0") },
        { 48, Number("##0.0E+0") },
        { 49, Number("@") },

        // Locale-specific built-in formats.
        // Indices 27–36 and 50–58 vary by locale (ja-jp, ko-kr, zh-tw, zh-cn) but are
        // always date/time formats. The ja-jp strings are stored as a representative
        // fallback; files from other locales may display differently via GetNumberFormatString()
        // but will still be correctly identified as date/time values.
        { 27, Date(@"[$-411]ge.m.d") },
        { 28, Date("[$-411]ggge\"年\"m\"月\"d\"日\"") },
        { 29, Date("[$-411]ggge\"年\"m\"月\"d\"日\"") },
        { 30, Date("m/d/yy") },
        { 31, Date("yyyy\"年\"m\"月\"d\"日\"") },
        { 32, Date("h\"時\"mm\"分\"") },
        { 33, Date("h\"時\"mm\"分\"ss\"秒\"") },
        { 34, Date("yyyy\"年\"m\"月\"") },
        { 35, Date("m\"月\"d\"日\"") },
        { 36, Date(@"[$-411]ge.m.d") },
        { 50, Date(@"[$-411]ge.m.d") },
        { 51, Date("[$-411]ggge\"年\"m\"月\"d\"日\"") },
        { 52, Date("yyyy\"年\"m\"月\"") },
        { 53, Date("m\"月\"d\"日\"") },
        { 54, Date("[$-411]ggge\"年\"m\"月\"d\"日\"") },
        { 55, Date("yyyy\"年\"m\"月\"") },
        { 56, Date("m\"月\"d\"日\"") },
        { 57, Date(@"[$-411]ge.m.d") },
        { 58, Date("[$-411]ggge\"年\"m\"月\"d\"日\"") },

        // 59–62 and 67–70 are Thai-digit number formats; 71–81 are Thai date/time formats.
        // ASCII equivalents are used so the format strings parse correctly.
        { 59, Number("0") },
        { 60, Number("0.00") },
        { 61, Number("#,##0") },
        { 62, Number("#,##0.00") },
        { 67, Number("0%") },
        { 68, Number("0.00%") },
        { 69, Number("# ?/?") },
        { 70, Number("# ??/??") },
        { 71, Date("d/m/yyyy") },
        { 72, Date("d-mmm-yy") },
        { 73, Date("d-mmm") },
        { 74, Date("mmm-yy") },
        { 75, Date("h:mm") },
        { 76, Date("h:mm:ss") },
        { 77, Date("d/m/yyyy h:mm") },
        { 78, Date("mm:ss") },
        { 79, Duration("[h]:mm:ss") },
        { 80, Date("mm:ss.0") },
        { 81, Date("d/m/yyyy") },
    };

    public static NumberFormatString? GetBuiltinNumberFormat(int numFmtId)
    {
        if (Formats.TryGetValue(numFmtId, out var result))
            return result;

        return null;
    }

    public static NumberFormatString? GetBuiltinNumberFormat(int numFmtId, IFormatProvider provider)
    {
        var dtf = DateTimeFormatInfo.GetInstance(provider);

        if (numFmtId == 14)
            return new NumberFormatString(DatePatternToExcel(dtf.ShortDatePattern), isDateTimeFormat: true, isTimeSpanFormat: false);

        if (numFmtId == 15 || numFmtId == 16 || numFmtId == 17)
        {
            var sep = EscapeExcelLiteral(dtf.DateSeparator);
            var s = numFmtId switch
            {
                15 => $"d{sep}mmm{sep}yy",
                16 => $"d{sep}mmm",
                _ => $"mmm{sep}yy",
            };
            return new NumberFormatString(s, isDateTimeFormat: true, isTimeSpanFormat: false);
        }

        if (numFmtId == 20)
        {
            var pattern = TimePatternToExcel(dtf.ShortTimePattern);
            if (!pattern.Contains("AM/PM") && !pattern.Contains("A/P"))
                return new NumberFormatString(pattern, isDateTimeFormat: true, isTimeSpanFormat: false);
        }

        if (numFmtId == 21)
        {
            var pattern = TimePatternToExcel(dtf.LongTimePattern);
            if (!pattern.Contains("AM/PM") && !pattern.Contains("A/P"))
                return new NumberFormatString(pattern, isDateTimeFormat: true, isTimeSpanFormat: false);
        }

        if (numFmtId == 22)
        {
            var date = DatePatternToExcel(dtf.ShortDatePattern);
            var time = TimePatternToExcel(dtf.ShortTimePattern);
            return new NumberFormatString($"{date} {time}", isDateTimeFormat: true, isTimeSpanFormat: false);
        }

        if (Formats.TryGetValue(numFmtId, out var result))
            return result;

        return null;
    }

    private static NumberFormatString Date(string s) => new(s, isDateTimeFormat: true, isTimeSpanFormat: false);

    private static NumberFormatString Duration(string s) => new(s, isDateTimeFormat: false, isTimeSpanFormat: true);

    private static NumberFormatString Number(string s) => new(s, isDateTimeFormat: false, isTimeSpanFormat: false);

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
