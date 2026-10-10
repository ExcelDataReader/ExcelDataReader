namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class CellRecord(int columnIndex, int xfIndex, DecodedCellValue value, CellError? error) : Record
{
    // Flat fields avoid the per-record padding of nested value/nullable structs on .NET Framework.
    private readonly object? _value = value.Value;
    private readonly CellValueKind _kind = value.Kind;
    private readonly uint _sharedStringIndex = value.SharedStringIndex;
    private readonly CellError _error = error.GetValueOrDefault();
    private readonly bool _hasError = error.HasValue;

    public int ColumnIndex { get; } = columnIndex;

    public int XfIndex { get; } = xfIndex;

    public DecodedCellValue Value => new(_value, _kind, _sharedStringIndex);

    public CellError? Error => _hasError ? _error : null;
}
