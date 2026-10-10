namespace ExcelDataReader.Core;

internal interface ISharedStringStore : ISharedStringSink, IDisposable
{
    int Count { get; }

    void Add(string value);

    void Reserve(int count);

    void Seal();

    string GetString(int index);
}
