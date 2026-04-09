#nullable enable

using System.Globalization;
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
  
    /// <summary>
    /// Gets or sets the culture to use for locale-dependent built-in number format indices.
    /// When null (the default), hardcoded format strings are used.
    /// </summary>
    public CultureInfo? Culture { get; set; }

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

    public NumberFormatString GetNumberFormatString(int numberFormatIndex)
    {
        if (Formats.TryGetValue(numberFormatIndex, out var numberFormat))
        {
            return numberFormat;
        }

        numberFormat = Culture != null
            ? BuiltinNumberFormat.GetBuiltinNumberFormat(numberFormatIndex, Culture)
#pragma warning disable CA1304 // Intentional: null Culture means use hardcoded formats for backward compatibility
            : BuiltinNumberFormat.GetBuiltinNumberFormat(numberFormatIndex);
#pragma warning restore CA1304
        if (numberFormat != null)
        {
            return numberFormat;
        }

        // Fall back to "General" if the number format index is invalid
        return GeneralNumberFormat;
    }
}
