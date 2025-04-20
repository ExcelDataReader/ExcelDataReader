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
        List<Section> sections = [];
        var isValid = true;
        while (true)
        {
            var section = Parser.ParseSection(tokenizer, out var syntaxError);

            if (syntaxError)
                isValid = false;

            if (section == null)
                break;

            sections.Add(section);
        }

        IsValid = isValid;
        FormatString = formatString;

        if (isValid)
        {
            Sections = sections;
            IsDateTimeFormat = GetFirstSection(SectionType.Date) != null;
            IsTimeSpanFormat = GetFirstSection(SectionType.Duration) != null;
        }
        else
        {
            Sections = [];
        }
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

    private IReadOnlyList<Section> Sections { get; }

    private Section GetFirstSection(SectionType type) => Sections.FirstOrDefault(section => section.Type == type);
}
