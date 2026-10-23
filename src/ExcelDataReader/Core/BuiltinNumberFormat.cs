using System.Globalization;
using System.Text;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core;

internal static class BuiltinNumberFormat
{
    private static NumberFormatString?[] Formats { get; } =
    [
        Number("General"),
        Number("0"),
        Number("0.00"),
        Number("#,##0"),
        Number("#,##0.00"),
        Number("\"$\"#,##0_);(\"$\"#,##0)"),
        Number("\"$\"#,##0_);[Red](\"$\"#,##0)"),
        Number("\"$\"#,##0.00_);(\"$\"#,##0.00)"),
        Number("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)"),
        Number("0%"),
        Number("0.00%"),
        Number("0.00E+00"),
        Number("# ?/?"),
        Number("# ??/??"),
        Date("d/m/yyyy"),
        Date("d-mmm-yy"),
        Date("d-mmm"),
        Date("mmm-yy"),
        Date("h:mm AM/PM"),
        Date("h:mm:ss AM/PM"),
        Date("h:mm"),
        Date("h:mm:ss"),
        Date("m/d/yy h:mm"),

        // 23..26 international/unused
        null, null, null, null,

        // Indices 27–36 and 50–58 vary by locale (ja-jp, ko-kr, zh-tw, zh-cn) but are
        // always date/time formats. The ja-jp strings are stored as a representative
        // fallback; files from other locales may display differently via GetNumberFormatString()
        // but will still be correctly identified as date/time values.
        Date(@"[$-411]ge.m.d"),
        Date("[$-411]ggge\"年\"m\"月\"d\"日\""),
        Date("[$-411]ggge\"年\"m\"月\"d\"日\""),
        Date("m/d/yy"),
        Date("yyyy\"年\"m\"月\"d\"日\""),
        Date("h\"時\"mm\"分\""),
        Date("h\"時\"mm\"分\"ss\"秒\""),
        Date("yyyy\"年\"m\"月\""),
        Date("m\"月\"d\"日\""),
        Date(@"[$-411]ge.m.d"),

        // 37..49
        Number("#,##0_);(#,##0)"),
        Number("#,##0_);[Red](#,##0)"),
        Number("#,##0.00_);(#,##0.00)"),
        Number("#,##0.00_);[Red](#,##0.00)"),
        Number("_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)"),
        Number("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"),
        Number("_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)"),
        Number("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"),
        Date("mm:ss"),
        Duration("[h]:mm:ss"),
        Date("mm:ss.0"),
        Number("##0.0E+0"),
        Number("@"),

        // 50..58 locale-specific date/time formats
        Date(@"[$-411]ge.m.d"),
        Date("[$-411]ggge\"年\"m\"月\"d\"日\""),
        Date("yyyy\"年\"m\"月\""),
        Date("m\"月\"d\"日\""),
        Date("[$-411]ggge\"年\"m\"月\"d\"日\""),
        Date("yyyy\"年\"m\"月\""),
        Date("m\"月\"d\"日\""),
        Date(@"[$-411]ge.m.d"),
        Date("[$-411]ggge\"年\"m\"月\"d\"日\""),

        // 59–62 and 67–70 are Thai-digit number formats; 71–81 are Thai date/time formats.
        // ASCII equivalents are used so the format strings parse correctly.
        Number("0"),
        Number("0.00"),
        Number("#,##0"),
        Number("#,##0.00"),

        // 63..66 unused
        null, null, null, null,
        Number("0%"),
        Number("0.00%"),
        Number("# ?/?"),
        Number("# ??/??"),
        Date("d/m/yyyy"),
        Date("d-mmm-yy"),
        Date("d-mmm"),
        Date("mmm-yy"),
        Date("h:mm"),
        Date("h:mm:ss"),
        Date("d/m/yyyy h:mm"),
        Date("mm:ss"),
        Duration("[h]:mm:ss"),
        Date("mm:ss.0"),
        Date("d/m/yyyy"),
    ];

    public static NumberFormatString? GetBuiltinNumberFormat(int numFmtId)
    {
        return (uint)numFmtId < (uint)Formats.Length ? Formats[numFmtId] : null;
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

        return (uint)numFmtId < (uint)Formats.Length ? Formats[numFmtId] : null;
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