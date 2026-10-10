using System.Globalization;
using System.Runtime.CompilerServices;
using System.Xml;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core;

/// <summary>
/// Common handling of extended formats (XF), number formats and decoded cell materialization.
/// </summary>
internal class CommonWorkbook
{
    /// <summary>
    /// Gets the dictionary of global number format strings. Always includes the built-in formats at their
    /// corresponding indices and any additional formats specified in the workbook file.
    /// </summary>
    public Dictionary<int, NumberFormatString> Formats { get; } = [];

    /// <summary>
    /// Gets the Cell XFs.
    /// </summary>
    public List<ExtendedFormat> ExtendedFormats { get; } = [];

    /// <summary>
    /// Gets the Cell Style XFs.
    /// </summary>
    public List<ExtendedFormat> CellStyleExtendedFormats { get; } = [];

    public bool SinglePassMode { get; set; }

    public ExtendedFormat GetEffectiveCellStyle(int xfIndex, int numberFormatFromCell)
    {
        if (xfIndex >= 0 && xfIndex < ExtendedFormats.Count)
        {
            return ExtendedFormats[xfIndex];
        }

        if (numberFormatFromCell == 0)
            return ExtendedFormat.Zero;

        return new ExtendedFormat(numberFormatFromCell);
    }

    /// <summary>
    /// Registers a number format string in the workbook's Formats dictionary.
    /// </summary>
    public void AddNumberFormat(int formatIndexInFile, string formatString)
    {
        if (!Formats.ContainsKey(formatIndexInFile))
            Formats.Add(formatIndexInFile, new NumberFormatString(formatString));
    }

    public object ConvertNumericValue(double value, int numberFormatIndex, bool date1904) =>
        ConvertNumericValueCore(value, numberFormatIndex, date1904);

    // BIFF2 integer cells must remain integers when no date conversion is possible.
    public object ConvertNumericValue(int value, int numberFormatIndex, bool date1904) =>
        ConvertNumericValueCore(value, numberFormatIndex, date1904);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Cell CreateCell(int columnIndex, DecodedCellValue decoded, ExtendedFormat style, CellError? error, bool date1904)
    {
        object? value = decoded.Value;
        value = decoded.Kind switch
        {
            CellValueKind.SharedString => ResolveSharedString(decoded.SharedStringIndex),
            CellValueKind.FormattedString => ConvertFormattedString((string)value!, style.NumberFormatIndex, date1904),
            _ when value is double or int => ConvertNumericValueCore(value, style.NumberFormatIndex, date1904),
            _ => value,
        };

        return new Cell(columnIndex, value, style, error);
    }

    public NumberFormatString? GetNumberFormatString(int numberFormatIndex, IFormatProvider? provider)
    {
        // User-defined formats (from the workbook file) take precedence.
        if (Formats.TryGetValue(numberFormatIndex, out var numberFormat))
            return numberFormat;

        // null provider means locale-independent built-in strings.
        if (provider == null)
        {
#pragma warning disable CA1305 // Intentional: null provider returns hardcoded locale-independent format strings
            return BuiltinNumberFormat.GetBuiltinNumberFormat(numberFormatIndex) ?? GeneralNumberFormatCache.Value;
#pragma warning restore CA1305
        }

        // For locale-sensitive built-in indices (14–17, 22) derive the pattern from the
        // provider; fall back to the hardcoded string for all other indices.
#pragma warning disable CA1305 // Intentional: fallback to hardcoded strings when no locale-specific override exists
        return BuiltinNumberFormat.GetBuiltinNumberFormat(numberFormatIndex, provider)
            ?? BuiltinNumberFormat.GetBuiltinNumberFormat(numberFormatIndex);
#pragma warning restore CA1305
    }

    protected virtual string? ResolveSharedString(uint index) => null;

    private static bool TryParseToTimeSpan(string s, out TimeSpan result)
    {
        if (!Helpers.StringStartsWith(s, 'P'))
            return TimeSpan.TryParse(s, out result);

        try
        {
            result = XmlConvert.ToTimeSpan(s);
            return true;
        }
        catch (FormatException)
        {
            result = TimeSpan.Zero;
            return false;
        }
    }

    private object ConvertFormattedString(string value, int numberFormatIndex, bool date1904)
    {
        var format = GetNumberFormatString(numberFormatIndex, null);
        if (format?.IsTimeSpanFormat == true && TryParseToTimeSpan(value, out var timeSpan))
            return timeSpan;

        if (format?.IsDateTimeFormat == true &&
            DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault, out DateTime date))
        {
            // NoCurrentDateDefault marks HH:mm:ss values with year 1; an explicit ISO date can also use that year.
            if (date.Date == DateTime.MinValue && value.TrimStart().IndexOf(':') == 2)
                return Helpers.ConvertFromOATime(date.TimeOfDay.TotalDays, date1904);

            return date;
        }

        return value;
    }

    private object ConvertNumericValueCore(object value, int numberFormatIndex, bool date1904)
    {
        var format = GetNumberFormatString(numberFormatIndex, null);
        if (format != null)
        {
            if (format.IsDateTimeFormat)
            {
                return value switch
                {
                    int integer => Helpers.ConvertFromOATime(integer, date1904),
                    double number => Helpers.ConvertFromOATime(number, date1904),
                    _ => value,
                };
            }

            if (format.IsTimeSpanFormat)
            {
                return value switch
                {
                    int integer => TimeSpan.FromDays(integer),
                    double number => TimeSpan.FromDays(number),
                    _ => value,
                };
            }
        }

        return value;
    }

    private static class GeneralNumberFormatCache
    {
        internal static NumberFormatString Value { get; } = new("General");
    }
}
