namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents a Shared String Table in BIFF8 format.
/// </summary>
internal sealed class XlsBiffSST : XlsBiffRecord
{
    private readonly XlsSSTReader _reader;
    private readonly ISharedStringStore _store;

    internal XlsBiffSST(byte[] bytes, ISharedStringStore store)
        : base(bytes)
    {
        _store = store;
        _reader = new XlsSSTReader();
        ReadSstStrings();
    }

    /// <summary>
    /// Gets the number of strings in SST.
    /// </summary>
    public uint Count => ReadUInt32(0x0);

    /// <summary>
    /// Gets the count of unique strings in SST.
    /// </summary>
    public uint UniqueCount => ReadUInt32(0x4);

    private uint StringCount => (uint)_store.Count;

    private int RemainingStringCount
    {
        get
        {
            uint remaining = UniqueCount > StringCount ? UniqueCount - StringCount : 0;
            return remaining > int.MaxValue ? int.MaxValue : (int)remaining;
        }
    }

    /// <summary>
    /// Parses strings out of a Continue record.
    /// </summary>
    public void ReadContinueStrings(XlsBiffContinue sstContinue)
    {
        if (StringCount == UniqueCount)
        {
            return;
        }

        _reader.ReadStringsFromContinue(sstContinue, _store, RemainingStringCount);
    }

    public void Flush() => _reader.Flush(_store);

    /// <summary>
    /// Parses strings out of this SST record.
    /// </summary>
    private void ReadSstStrings()
    {
        if (StringCount == UniqueCount)
        {
            return;
        }

        _reader.ReadStringsFromSST(this, _store, RemainingStringCount);
    }
}