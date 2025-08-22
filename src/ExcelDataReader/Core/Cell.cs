#nullable enable

namespace ExcelDataReader.Core;

/// <summary>
/// Represents a cell.
/// </summary>
/// <param name="ColumnIndex">The zero-based column index.</param>
/// <param name="Value">The value.</param>
/// <param name="Hyperlink">The URL if cell is a hyperlink</param>
/// <param name="EffectiveStyle">
/// The effective style on the cell. The effective style is determined from 
/// the Cell XF, with optional overrides from a Cell Style XF.
/// </param>
/// <param name="Error">Cell error -or- <s langword="null"/>.</param>
internal sealed record Cell(int ColumnIndex, object? Value, string? Hyperlink, ExtendedFormat EffectiveStyle, CellError? Error);
