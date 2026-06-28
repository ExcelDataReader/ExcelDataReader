using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.CompoundFormat;

internal sealed class CompoundStream : Stream
{
    private readonly byte[] _sectorBuffer;

    private int _sectorBufferValidLength;

    public CompoundStream(CompoundDocument document, Stream baseStream, List<uint> sectorChain, int length, bool leaveOpen)
    {
        Document = document;
        BaseStream = baseStream;
        IsMini = false;
        LeaveOpen = leaveOpen;
        Length = length;
        SectorChain = sectorChain;
        _sectorBuffer = new byte[Document.Header.SectorSize];
        ReadSector();
    }

    public CompoundStream(CompoundDocument document, Stream baseStream, uint baseSector, int length, bool isMini, bool leaveOpen)
    {
        Document = document;
        BaseStream = baseStream;
        IsMini = isMini;
        Length = length;
        LeaveOpen = leaveOpen;

        if (IsMini)
        {
            SectorChain = CompoundDocument.GetSectorChain(baseSector, Document.MiniSectorTable);
            RootSectorChain = CompoundDocument.GetSectorChain(Document.RootEntry.StreamFirstSector, Document.SectorTable);
            _sectorBuffer = new byte[Document.Header.MiniSectorSize];
        }
        else
        {
            SectorChain = CompoundDocument.GetSectorChain(baseSector, Document.SectorTable);
            _sectorBuffer = new byte[Document.Header.SectorSize];
        }

        ReadSector();
    }

    public List<uint> SectorChain { get; }

    public List<uint>? RootSectorChain { get; }

    public override bool CanRead => true;

    public override bool CanSeek => false;

    public override bool CanWrite => false;

    public override long Length { get; }

    public override long Position { get => Offset - _sectorBufferValidLength + SectorOffset; set => Seek(value, SeekOrigin.Begin); }

    private Stream? BaseStream { get; set; }

    private CompoundDocument Document { get; }

    private bool IsMini { get; }

    private bool LeaveOpen { get; }

    private int SectorChainOffset { get; set; }

    private int Offset { get; set; }

    private int SectorOffset { get; set; }

    public override void Flush()
    {
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        int index = 0;
        while (index < count && Position < Length)
        {
            if (SectorOffset == _sectorBufferValidLength)
            {
                ReadSector();
                SectorOffset = 0;
            }

            var chunkSize = Math.Min(count - index, _sectorBufferValidLength - SectorOffset);
            Array.Copy(_sectorBuffer, SectorOffset, buffer, offset + index, chunkSize);
            index += chunkSize;
            SectorOffset += chunkSize;
        }

        return index;
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        var sectorSize = IsMini ? Document.Header.MiniSectorSize : Document.Header.SectorSize;
        switch (origin)
        {
            case SeekOrigin.Begin:
                SectorChainOffset = (int)(offset / sectorSize);
                Offset = SectorChainOffset * sectorSize;
                SectorOffset = (int)(offset % sectorSize);
                if (Offset < Length)
                    ReadSector();
                return Position;
            case SeekOrigin.Current:
                return Seek(Position + offset, SeekOrigin.Begin);
            case SeekOrigin.End:
                return Seek(Length + offset, SeekOrigin.Begin);
            default:
                return Offset;
        }
    }

    public override void SetLength(long value)
    {
        throw new NotImplementedException();
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        throw new NotImplementedException();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing && !LeaveOpen)
        {
            BaseStream?.Dispose();
            BaseStream = null;
        }

        base.Dispose(disposing);
    }

    private void ReadSector()
    {
        if (IsMini)
        {
            ReadMiniSector();
        }
        else
        {
            ReadRegularSector();
        }
    }

    private void ReadMiniSector()
    {
        var baseStream = BaseStream ?? throw new ObjectDisposedException(nameof(CompoundStream));

        if (RootSectorChain == null)
        {
            throw new InvalidOperationException("Mini stream sector chain is not initialized.");
        }

        var sector = SectorChain[SectorChainOffset];
        var miniStreamOffset = (int)Document.GetMiniSectorOffset(sector);

        var rootSectorIndex = miniStreamOffset / Document.Header.SectorSize;
        if (rootSectorIndex >= RootSectorChain.Count)
        {
            throw new CompoundDocumentException(Errors.ErrorEndOfFile);
        }

        var rootSector = RootSectorChain[rootSectorIndex];
        var rootOffset = miniStreamOffset % Document.Header.SectorSize;

        baseStream.Seek(Document.GetSectorOffset(rootSector) + rootOffset, SeekOrigin.Begin);

        var chunkSize = (int)Math.Min(Length - Offset, Document.Header.MiniSectorSize);
        if (baseStream.ReadAtLeast(_sectorBuffer, 0, chunkSize) < chunkSize)
        {
            throw new CompoundDocumentException(Errors.ErrorEndOfFile);
        }

        _sectorBufferValidLength = chunkSize;
        Offset += chunkSize;
        SectorChainOffset++;
    }

    private void ReadRegularSector()
    {
        var baseStream = BaseStream ?? throw new ObjectDisposedException(nameof(CompoundStream));

        var sector = SectorChain[SectorChainOffset];
        baseStream.Seek(Document.GetSectorOffset(sector), SeekOrigin.Begin);

        var chunkSize = (int)Math.Min(Length - Offset, Document.Header.SectorSize);
        if (baseStream.ReadAtLeast(_sectorBuffer, 0, chunkSize) < chunkSize)
        {
            throw new CompoundDocumentException(Errors.ErrorEndOfFile);
        }

        _sectorBufferValidLength = chunkSize;
        Offset += chunkSize;
        SectorChainOffset++;
    }
}
