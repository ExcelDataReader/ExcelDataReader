namespace ExcelDataReader.Sample;

internal sealed class SheetGridRow
{
    private readonly object?[] _values;

    public SheetGridRow(object?[] values)
    {
        _values = values;
    }

    public object? this[int index] => _values[index];
}
