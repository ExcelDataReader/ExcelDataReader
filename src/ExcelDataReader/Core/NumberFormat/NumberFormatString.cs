namespace ExcelDataReader.Core.NumberFormat;

internal enum ConditionOperator
{
    LessThan,
    LessThanOrEqual,
    Equal,
    NotEqual,
    GreaterThanOrEqual,
    GreaterThan,
}

internal readonly record struct NumberFormatSection(SectionType? Type, NumberFormatCondition? Condition);

internal readonly record struct NumberFormatCondition(ConditionOperator Operator, double Value)
{
    public bool Matches(double value) => Operator switch
    {
        ConditionOperator.LessThan => value < Value,
        ConditionOperator.LessThanOrEqual => value <= Value,
        ConditionOperator.Equal => value == Value,
        ConditionOperator.NotEqual => value != Value,
        ConditionOperator.GreaterThanOrEqual => value >= Value,
        ConditionOperator.GreaterThan => value > Value,
        _ => false,
    };
}

/// <summary>
/// Parse ECMA-376 number format strings from Excel and other spreadsheet softwares.
/// </summary>
public class NumberFormatString
{
    private readonly List<NumberFormatSection> _sections = [];

    /// <summary>
    /// Initializes a new instance of the <see cref="NumberFormatString"/> class.
    /// </summary>
    /// <param name="formatString">The number format string.</param>
    public NumberFormatString(string formatString)
    {
        Tokenizer tokenizer = new(formatString);
        List<Token> tokens = [];
        var isValid = true;
        bool isDateTimeFormat = false;
        bool isTimeSpanFormat = false;
        bool aggregateClassification = true;
        bool validateSections = true;
        bool moreSections;
        do
        {
            SectionType? section;
            bool syntaxError;
            NumberFormatCondition? condition;
            try
            {
                section = Parser.ParseSection(tokenizer, tokens, out syntaxError, out moreSections, out condition);
            }
            catch (Exception exception) when (!validateSections && exception is (FormatException or OverflowException or ArgumentOutOfRangeException))
            {
                _sections.Add(new NumberFormatSection(null, null));
                break;
            }

            if (syntaxError && validateSections)
                isValid = false;

            _sections.Add(new NumberFormatSection(section, condition));
            if (section == null)
            {
                aggregateClassification = false;
                validateSections = false;
            }
            else if (aggregateClassification)
            {
                isDateTimeFormat |= section == SectionType.Date;
                isTimeSpanFormat |= section == SectionType.Duration;
            }
        }
        while (moreSections);

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

        Tokenizer tokenizer = new(formatString);
        var tokens = new List<Token>();
        var section = Parser.ParseSection(tokenizer, tokens, out _, out _, out _);
        var type = isDateTimeFormat ? SectionType.Date : isTimeSpanFormat ? SectionType.Duration : section ?? SectionType.Number;
        _sections.Add(new NumberFormatSection(type, null));
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

    internal NumberFormatSection? GetSection(double value)
    {
        if (!IsValid)
            return null;

        int conditionalCount = 0;
        int conditionalIndex = -1;
        int conditionLimit = Math.Min(2, _sections.Count);
        for (int i = 0; i < conditionLimit; i++)
        {
            if (_sections[i].Condition != null)
            {
                conditionalCount++;
                conditionalIndex = i;
            }
        }

        if (conditionalCount > 0)
        {
            for (int i = 0; i < conditionLimit; i++)
            {
                if (_sections[i].Condition is { } condition && condition.Matches(value))
                    return _sections[i];
            }

            if (conditionalCount > 1)
                return _sections.Count > 2 ? _sections[2] : null;

            int fallbackIndex = conditionalIndex == 0 ? 1 : 0;
            return _sections.Count > fallbackIndex ? _sections[fallbackIndex] : null;
        }

        if (double.IsNaN(value))
            return null;

        int sectionIndex = _sections.Count switch
        {
            0 => -1,
            1 => 0,
            2 => value >= 0 ? 0 : 1,
            _ => value > 0 ? 0 : value < 0 ? 1 : 2,
        };
        return sectionIndex >= 0 ? _sections[sectionIndex] : null;
    }
}
