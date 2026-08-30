using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents a Shared String Table in BIFF8 format.
/// </summary>
internal sealed class XlsBiffSST : XlsBiffRecord
{
    private readonly XlsSSTReader _reader = new();
    private string?[]? _materializedStrings;
    private List<IXlsString?> _strings = [];

    internal XlsBiffSST(byte[] bytes)
        : base(bytes)
    {
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

    /// <summary>
    /// Parses strings out of a Continue record.
    /// </summary>
    public void ReadContinueStrings(XlsBiffContinue sstContinue)
    {
        if (_strings.Count == UniqueCount)
        {
            return;
        }

        foreach (var str in _reader.ReadStringsFromContinue(sstContinue))
        {
            _strings.Add(str);

            if (_strings.Count == UniqueCount)
            {
                break;
            }
        }
    }

    public void Flush()
    {
        var str = _reader.Flush();
        if (str != null)
        {
            _strings.Add(str);
        }

        _materializedStrings = new string[_strings.Count];
    }

    /// <summary>
    /// Returns string at specified index, caching it on first access and releasing the raw byte buffer.
    /// </summary>
    /// <param name="sstIndex">Index of string to get.</param>
    /// <param name="encoding">Workbook encoding.</param>
    /// <returns>string value if it was found, null otherwise.</returns>
    public string? GetString(uint sstIndex, Encoding encoding)
    {
        if (_materializedStrings == null)
        {
            if (sstIndex < _strings.Count)
                return _strings[(int)sstIndex]?.GetValue(encoding);
            return null;
        }

        if (sstIndex >= (uint)_materializedStrings.Length)
            return null;

        var cached = _materializedStrings[sstIndex];
        if (cached != null)
            return cached;

        var s = _strings[(int)sstIndex]?.GetValue(encoding);
        _materializedStrings[sstIndex] = s;
        _strings[(int)sstIndex] = null;
        return s;
    }

    /// <summary>
    /// Parses strings out of this SST record.
    /// </summary>
    private void ReadSstStrings()
    {
        if (_strings.Count == UniqueCount)
        {
            return;
        }

        foreach (var str in _reader.ReadStringsFromSST(this))
        {
            _strings.Add(str);

            if (_strings.Count == UniqueCount)
            {
                break;
            }
        }
    }
}