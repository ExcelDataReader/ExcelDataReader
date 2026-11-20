namespace ExcelDataReader.Core.BinaryFormat;

internal struct BiffFrameBuffer(int minimumCapacity)
{
    private byte[]? _buffer;

    public byte[] GetBuffer(int size)
    {
        if (_buffer == null || _buffer.Length < size)
            _buffer = new byte[Math.Max(minimumCapacity, size)];

        return _buffer;
    }

    public void Clear() => _buffer = null;
}