namespace ExcelDataReader.Core;

internal static class DateTimeHelper
{
    // All OA dates must be greater than (not >=) OADateMinAsDouble
    public const double OADateMinAsDouble = -657435.0;

    // All OA dates must be less than (not <=) OADateMaxAsDouble
    public const double OADateMaxAsDouble = 2958466.0;

    // Number of days in a non-leap year
    private const int DaysPerYear = 365;

    // Number of days in 4 years
    private const int DaysPer4Years = DaysPerYear * 4 + 1;

    // Number of days in 100 years
    private const int DaysPer100Years = DaysPer4Years * 25 - 1;

    // Number of days in 400 years
    private const int DaysPer400Years = DaysPer100Years * 4 + 1;

    // Number of days from 1/1/0001 to 12/30/1899
    private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;

    // Number of days from 1/1/0001 to 12/31/9999
    private const int DaysTo10000 = DaysPer400Years * 25 - 366;

    private const long MaxTicks = (long)DaysTo10000 * TimeSpan.TicksPerDay;

    private const long DoubleDateOffset = DaysTo1899 * TimeSpan.TicksPerDay;

    public static TimeSpan DurationFromDays(double value) => new(DaysToTicks(value));

    public static DateTime FromOADate(double value, int dayOffset)
    {
        long ticks = DaysToTicks(value) + (dayOffset * TimeSpan.TicksPerDay);

        // Negative OA dates have a negative day component but a positive time of day.
        if (ticks < 0)
            ticks -= ticks % TimeSpan.TicksPerDay * 2;

        ticks += DoubleDateOffset;
        if (ticks is < 0 or >= MaxTicks)
            throw new ArgumentException("OA Date out of range", nameof(value));

        return new DateTime(ticks, DateTimeKind.Unspecified);
    }

    private static long DaysToTicks(double value)
    {
        if (double.IsNaN(value))
            throw new ArgumentException("Value cannot be NaN.", nameof(value));
        if (value > TimeSpan.MaxValue.TotalDays || value < TimeSpan.MinValue.TotalDays)
            throw new OverflowException("Value is outside the range of a TimeSpan.");

        // Keep whole days separate so large serials do not lose fractional ticks during multiplication.
        long days = (long)value;
        long fraction = (long)Math.Round((value - days) * TimeSpan.TicksPerDay, MidpointRounding.AwayFromZero);
        return checked((days * TimeSpan.TicksPerDay) + fraction);
    }
}