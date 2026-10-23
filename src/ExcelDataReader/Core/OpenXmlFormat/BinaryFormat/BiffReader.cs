using System.Text;

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;

internal abstract class BiffReader(Stream stream) : IDisposable
{
    private readonly byte[] _buffer = new byte[512];

    private Stream Stream { get; } = stream ?? throw new ArgumentNullException(nameof(stream));

    public virtual void Dispose() => Stream.Dispose();

    protected static uint GetDWord(byte[] buffer, uint offset)
    {
        uint result = (uint)buffer[offset + 3] << 24;
        result += (uint)buffer[offset + 2] << 16;
        result += (uint)buffer[offset + 1] << 8;
        result += buffer[offset];
        return result;
    }

    protected static int GetInt32(byte[] buffer, uint offset)
    {
        int result = buffer[offset + 3] << 24;
        result += buffer[offset + 2] << 16;
        result += buffer[offset + 1] << 8;
        result += buffer[offset];
        return result;
    }

    protected static ushort GetWord(byte[] buffer, uint offset)
    {
        ushort result = (ushort)(buffer[offset + 1] << 8);
        result += buffer[offset];
        return result;
    }

    protected static string GetString(byte[] buffer, uint offset, uint length)
    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        return Encoding.Unicode.GetString(buffer.AsSpan((int)offset, (int)(length * 2)));
#else
        return Encoding.Unicode.GetString(buffer, (int)offset, (int)(length * 2));
#endif
    }

    protected static string? GetNullableString(byte[] buffer, ref uint offset)
    {
        var length = GetDWord(buffer, offset);
        offset += 4;
        if (length == uint.MaxValue)
            return null;
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        string result = Encoding.Unicode.GetString(buffer.AsSpan((int)offset, (int)(length * 2)));
#else
        string result = Encoding.Unicode.GetString(buffer, (int)offset, (int)(length * 2));
#endif
        offset += length * 2;
        return result;
    }

    protected static double GetRkNumber(byte[] buffer, uint offset)
    {
        double result;

        byte flags = buffer[offset];

        if ((flags & 0x02) != 0)
        {
            result = GetInt32(buffer, offset) >> 2;
        }
        else
        {
            result = BitConverter.Int64BitsToDouble((GetDWord(buffer, offset) & -4) << 32);
        }

        if ((flags & 0x01) != 0)
        {
            result /= 100;
        }

        return result;
    }

    protected static double GetDouble(byte[] buffer, uint offset)
    {
        uint num = GetDWord(buffer, offset);
        uint num2 = GetDWord(buffer, offset + 4);
        long num3 = ((long)num2 << 32) | num;
        return BitConverter.Int64BitsToDouble(num3);
    }

    // A successful frame is borrowed until ReleaseFrame on pooled targets.
    // Keep non-pooled framing inlineable without unnecessary cleanup regions.
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    protected bool TryReadFrame(out byte[] buffer, out uint recordId, out uint recordLength)
    {
        buffer = _buffer;
        recordLength = 0;
        if (!TryReadVariableValue(out recordId) ||
            !TryReadVariableValue(out recordLength))
            return false;

        buffer = GetFrameBuffer(recordLength);
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        try
        {
            if (Stream.ReadAtLeast(buffer, 0, (int)recordLength) != recordLength)
            {
                ReleaseFrame(buffer);
                return false;
            }

            return true;
        }
        catch
        {
            ReleaseFrame(buffer);
            throw;
        }
#else
        return Stream.ReadAtLeast(buffer, 0, (int)recordLength) == recordLength;
#endif
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1822:Mark members as static", Justification = "Instance state is only used on pooled targets.")]
    protected void ReleaseFrame(byte[] buffer)
    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        if (ReferenceEquals(buffer, _buffer))
            return;

        System.Buffers.ArrayPool<byte>.Shared.Return(buffer);
#endif
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private byte[] GetFrameBuffer(uint recordLength)
    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        return recordLength >= _buffer.Length
            ? System.Buffers.ArrayPool<byte>.Shared.Rent((int)recordLength)
            : _buffer;
#else
        return recordLength < _buffer.Length ? _buffer : new byte[recordLength];
#endif
    }

    private bool TryReadVariableValue(out uint value)
    {
        value = 0;

        if (Stream.Read(_buffer, 0, 1) == 0)
            return false;

        byte b1 = _buffer[0];
        value = (uint)(b1 & 0x7F);

        if ((b1 & 0x80) == 0)
            return true;

        if (Stream.Read(_buffer, 0, 1) == 0)
            return false;
        byte b2 = _buffer[0];
        value = ((uint)(b2 & 0x7F) << 7) | value;

        if ((b2 & 0x80) == 0)
            return true;

        if (Stream.Read(_buffer, 0, 1) == 0)
            return false;
        byte b3 = _buffer[0];
        value = ((uint)(b3 & 0x7F) << 14) | value;

        if ((b3 & 0x80) == 0)
            return true;

        if (Stream.Read(_buffer, 0, 1) == 0)
            return false;
        byte b4 = _buffer[0];
        value = ((uint)(b4 & 0x7F) << 21) | value;

        return true;
    }
}