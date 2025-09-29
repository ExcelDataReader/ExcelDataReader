#nullable enable

namespace ExcelDataReader.Core;

/// <summary>
/// Represents a row.
/// </summary>
/// <param name="RowIndex">The zero-based row index.</param>
/// <param name="Height">The height of this row in points, zero if hidden or collapsed.</param>
/// <param name="Cells">The cells.</param>
internal sealed record Row(int RowIndex, double Height, List<Cell> Cells);