using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core;

/// <summary>
/// Common handling of extended formats (XF) and mappings between file-based and global number format indices.
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

    private NumberFormatString GeneralNumberFormat { get; } = new("General");

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

    public NumberFormatString? GetNumberFormatString(int numberFormatIndex, IFormatProvider? provider)
    {
        // User-defined formats (from the workbook file) take precedence.
        if (Formats.TryGetValue(numberFormatIndex, out var numberFormat))
            return numberFormat;

        // null provider means locale-independent built-in strings.
        if (provider == null)
        {
#pragma warning disable CA1305 // Intentional: null provider returns hardcoded locale-independent format strings
            return BuiltinNumberFormat.GetBuiltinNumberFormat(numberFormatIndex) ?? GeneralNumberFormat;
#pragma warning restore CA1305
        }

        // For locale-sensitive built-in indices (14–17, 22) derive the pattern from the
        // provider; fall back to the hardcoded string for all other indices.
#pragma warning disable CA1305 // Intentional: fallback to hardcoded strings when no locale-specific override exists
        return BuiltinNumberFormat.GetBuiltinNumberFormat(numberFormatIndex, provider)
            ?? BuiltinNumberFormat.GetBuiltinNumberFormat(numberFormatIndex);
#pragma warning restore CA1305
    }
}
