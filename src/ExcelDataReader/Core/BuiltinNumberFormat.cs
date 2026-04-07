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
