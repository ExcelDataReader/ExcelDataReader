using System.IO.MemoryMappedFiles;

namespace ExcelDataReader.Core;

internal sealed class MappedSharedStringReader : IDisposable
{
    internal const long ResidentBytes = 2048;
    private readonly MemoryMappedFile _mapping;
    private readonly long _length;
    private readonly int _viewSize;
    private MemoryMappedViewAccessor? _view;
    private long _viewOffset;
    private long _viewLength;
    private bool _disposed;

    public MappedSharedStringReader(FileStream file, long length, int viewSize = 0)
    {
        if (file.Length != length)
            throw new EndOfStreamException("Unexpected length of shared string storage.");
        _length = length;
        _viewSize = viewSize == 0 ? (IntPtr.Size == 4 ? 8 : 64) * 1024 * 1024 : viewSize;
        if (_viewSize <= 0 || _viewSize % 65536 != 0)
            throw new ArgumentOutOfRangeException(nameof(viewSize));
        _mapping = MemoryMappedFile.CreateFromFile(file, null, 0, MemoryMappedFileAccess.Read, HandleInheritability.None, leaveOpen: true);
    }

    public bool TryReadExactly(long offset, byte[] buffer, int count, bool adjacent)
    {
#if NET8_0_OR_GREATER
        ObjectDisposedException.ThrowIf(_disposed, this);
#else
        if (_disposed)
            throw new ObjectDisposedException(nameof(MappedSharedStringReader));
#endif
        if (offset < 0 || count < 0 || count > buffer.Length || offset > _length - count)
            throw new EndOfStreamException("Unexpected end of shared string storage.");

        // Unrelated misses outside the current view use buffered I/O, avoiding remapping thrash.
        if (_view != null && !adjacent && (offset < _viewOffset || offset + count > _viewOffset + _viewLength))
            return false;

        int copied = 0;
        while (copied < count)
        {
            long position = offset + copied;
            if (_view == null || position < _viewOffset || position >= _viewOffset + _viewLength)
            {
                // Bounded aligned views also work when the file exceeds a 32-bit address space.
                _view?.Dispose();
                _view = null;
                _viewOffset = position / _viewSize * _viewSize;
                _viewLength = Math.Min(_viewSize, _length - _viewOffset);
                _view = _mapping.CreateViewAccessor(_viewOffset, _viewLength, MemoryMappedFileAccess.Read);
            }

            int read = (int)Math.Min(count - copied, _viewOffset + _viewLength - position);
            if (_view.ReadArray(position - _viewOffset, buffer, copied, read) != read)
                throw new EndOfStreamException("Unexpected end of shared string storage.");
            copied += read;
        }

        return true;
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        try
        {
            ResourceCleanup.DisposeAll(null, _view, _mapping);
        }
        finally
        {
            _view = null;
        }
    }
}
