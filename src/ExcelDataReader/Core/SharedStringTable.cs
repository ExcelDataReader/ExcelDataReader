using System.Text;

namespace ExcelDataReader.Core;

/// <summary>
/// In-memory shared string table receiving UTF-16 strings.
/// </summary>
internal sealed class SharedStringTable : List<string>, ISharedStringStore
{
    public void AddUtf16(byte[] buffer, int offset, int characterCount) =>
        Add(Encoding.Unicode.GetString(buffer, offset, characterCount * 2));

    public void Reserve(int count)
    {
        if (Capacity < count)
            Capacity = count;
    }

    public string GetString(int index) => this[index];

    public void Seal()
    {
    }

    public void Dispose()
    {
    }
}
