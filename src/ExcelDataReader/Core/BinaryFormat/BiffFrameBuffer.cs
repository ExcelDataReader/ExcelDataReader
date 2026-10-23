namespace ExcelDataReader.Core.BinaryFormat;

internal struct BiffFrameBuffer
{
    private readonly int _minimumCapacity;
    private byte[]? _buffer;

    public BiffFrameBuffer(int minimumCapacity)
    {
        _minimumCapacity = minimumCapacity;
    }

    public byte[] GetBuffer(int size)
    {
        if (_buffer == null || _buffer.Length < size)
            _buffer = new byte[Math.Max(_minimumCapacity, size)];

        return _buffer;
    }

    public void Clear() => _buffer = null;
}