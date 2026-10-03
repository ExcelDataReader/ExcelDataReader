namespace ExcelDataReader.Core.NumberFormat;

/// <summary>
/// Parse ECMA-376 number format strings from Excel and other spreadsheet softwares.
/// </summary>
public class NumberFormatString
{
    /// <summary>
    /// Initializes a new instance of the <see cref="NumberFormatString"/> class.
    /// </summary>
    /// <param name="formatString">The number format string.</param>
    public NumberFormatString(string formatString)
    {
        Tokenizer tokenizer = new(formatString);
        var isValid = true;
        bool isDateTimeFormat = false;
        bool isTimeSpanFormat = false;
        while (true)
        {
            var section = Parser.ParseSection(tokenizer, out var syntaxError);

            if (syntaxError)
                isValid = false;

            if (section == null)
                break;

            isDateTimeFormat |= section.Type == SectionType.Date;
            isTimeSpanFormat |= section.Type == SectionType.Duration;
        }

        IsValid = isValid;
        FormatString = formatString;

        if (isValid)
        {
            IsDateTimeFormat = isDateTimeFormat;
            IsTimeSpanFormat = isTimeSpanFormat;
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="NumberFormatString"/> class with
    /// pre-known classification, bypassing parsing. Used for built-in format indices whose
    /// <see cref="IsDateTimeFormat"/> and <see cref="IsTimeSpanFormat"/> values are fixed by spec.
    /// </summary>
    internal NumberFormatString(string formatString, bool isDateTimeFormat, bool isTimeSpanFormat)
    {
        IsValid = true;
        FormatString = formatString;
        IsDateTimeFormat = isDateTimeFormat;
        IsTimeSpanFormat = isTimeSpanFormat;
    }

    /// <summary>
    /// Gets a value indicating whether the number format string is valid.
    /// </summary>
    public bool IsValid { get; }

    /// <summary>
    /// Gets the number format string.
    /// </summary>
    public string FormatString { get; }

    /// <summary>
    /// Gets a value indicating whether the format represents a DateTime.
    /// </summary>
    public bool IsDateTimeFormat { get; }

    /// <summary>
    /// Gets a value indicating whether the format represents a TimeSpan.
    /// </summary>
    public bool IsTimeSpanFormat { get; }
}
