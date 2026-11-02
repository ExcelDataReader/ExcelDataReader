using System.Collections;
using System.Data;

namespace ExcelDataReader.Sample;

internal sealed class SheetRows(DataRowCollection rows) : IList
{
    public int Count => rows.Count;

    public bool IsReadOnly => true;

    public bool IsFixedSize => true;

    public bool IsSynchronized => false;

    public object SyncRoot => ((ICollection)rows).SyncRoot;

    public object? this[int index]
    {
        get => rows[index];
        set => throw new NotSupportedException();
    }

    public IEnumerator GetEnumerator() => rows.GetEnumerator();

    public void CopyTo(Array array, int index) => rows.CopyTo(array, index);

    public bool Contains(object? value) => IndexOf(value) >= 0;

    public int IndexOf(object? value) => value is DataRow row ? rows.IndexOf(row) : -1;

    public int Add(object? value) => throw new NotSupportedException();

    public void Clear() => throw new NotSupportedException();

    public void Insert(int index, object? value) => throw new NotSupportedException();

    public void Remove(object? value) => throw new NotSupportedException();

    public void RemoveAt(int index) => throw new NotSupportedException();
}
