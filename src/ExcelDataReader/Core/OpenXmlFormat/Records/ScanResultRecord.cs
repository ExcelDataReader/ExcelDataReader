namespace ExcelDataReader.Core.OpenXmlFormat.Records;

/// <summary>
/// Produced by worksheet readers in preparing (pre-scan) mode instead of individual
/// RowHeaderRecord and CellRecord instances. Contains the maximum row and column indices
/// found in the sheet data, allowing the worksheet constructor to determine FieldCount
/// and RowCount without allocating per-cell objects.
/// </summary>
internal sealed class ScanResultRecord(int maxRowIndex, int maxColumnIndex) : Record
{
    public int MaxRowIndex { get; } = maxRowIndex;

    public int MaxColumnIndex { get; } = maxColumnIndex;
}
