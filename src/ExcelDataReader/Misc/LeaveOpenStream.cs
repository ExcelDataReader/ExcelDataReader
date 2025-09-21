#nullable enable

namespace ExcelDataReader.Misc;

internal sealed class LeaveOpenStream(Stream baseStream) : Stream
{
    public override bool CanRead => BaseStream.CanRead;

    public override bool CanSeek => BaseStream.CanSeek;

    public override bool CanWrite => BaseStream.CanWrite;

    public override long Length => BaseStream.Length;

    public override long Position { get => BaseStream.Position; set => BaseStream.Position = value; }

    private Stream BaseStream { get; } = baseStream;

    public override void Flush() => BaseStream.Flush();

    public override int Read(byte[] buffer, int offset, int count) => BaseStream.Read(buffer, offset, count);

    public override long Seek(long offset, SeekOrigin origin) => BaseStream.Seek(offset, origin);

    public override void SetLength(long value) => BaseStream.SetLength(value);

    public override void Write(byte[] buffer, int offset, int count) => BaseStream.Write(buffer, offset, count);
}
