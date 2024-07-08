namespace ExcelDataReader.Core.NumberFormat;

internal sealed class Section
{
    public required SectionType Type { get; init; }

    public required Color Color { get; init; }

    public required Condition Condition { get; init; }

    public required ExponentialSection Exponential { get; init; }

    public required FractionSection Fraction { get; init; }

    public required DecimalSection Number { get; init; }

    public required List<string> GeneralTextDateDurationParts { get; init; }
}