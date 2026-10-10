namespace ExcelDataReader.Core;

internal enum CellValueKind : byte
{
    Literal,
    SharedString,
    FormattedString,
}

/// <summary>
/// An owned cell value decoded by a format reader, before workbook-dependent conversion.
/// Shared string references are tagged explicitly so BIFF2 integers remain numeric values.
/// Formatted strings opt into date/duration parsing; legacy BIFF strings remain literal.
/// </summary>
internal readonly record struct DecodedCellValue(object? Value, CellValueKind Kind = CellValueKind.Literal, uint SharedStringIndex = 0)
{
    public static DecodedCellValue SharedString(uint index) => new(null, CellValueKind.SharedString, index);
}
