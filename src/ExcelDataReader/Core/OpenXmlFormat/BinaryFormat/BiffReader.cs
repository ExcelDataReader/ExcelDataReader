using System.Text;
using ExcelDataReader.Core.BinaryFormat;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;

internal abstract class BiffReader(Stream stream) : IDisposable
{
    private BiffFrameBuffer _frameBuffer = new(512);

    private Stream Stream { get; } = stream ?? throw new ArgumentNullException(nameof(stream));

    public virtual void Dispose()
    {
        _frameBuffer.Clear();
        Stream.Dispose();
    }

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
        string result = GetString(buffer, offset, length);
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

    // The returned frame borrows the reusable buffer until the next call.
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    protected bool TryReadFrame(out byte[] buffer, out uint recordId, out uint recordLength)
    {
        buffer = _frameBuffer.GetBuffer(1);
        recordLength = 0;
        if (!TryReadVariableValue(buffer, out recordId))
            return false;
        if (!TryReadVariableValue(buffer, out recordLength))
            throw new ExcelReaderException(Errors.ErrorBiffRecordSizeInvalid);

        buffer = _frameBuffer.GetBuffer((int)recordLength);
        if (Stream.ReadAtLeast(buffer, 0, (int)recordLength, (int)recordLength, throwOnEndOfStream: false) != recordLength)
            throw new ExcelReaderException(Errors.ErrorBiffRecordSizeInvalid);

        return true;
    }

    private bool TryReadVariableValue(byte[] buffer, out uint value)
    {
        value = 0;

        if (Stream.Read(buffer, 0, 1) == 0)
            return false;

        byte b1 = buffer[0];
        value = (uint)(b1 & 0x7F);

        if ((b1 & 0x80) == 0)
            return true;

        if (Stream.Read(buffer, 0, 1) == 0)
            throw new ExcelReaderException(Errors.ErrorBiffRecordSizeInvalid);
        byte b2 = buffer[0];
        value = ((uint)(b2 & 0x7F) << 7) | value;

        if ((b2 & 0x80) == 0)
            return true;

        if (Stream.Read(buffer, 0, 1) == 0)
            throw new ExcelReaderException(Errors.ErrorBiffRecordSizeInvalid);
        byte b3 = buffer[0];
        value = ((uint)(b3 & 0x7F) << 14) | value;

        if ((b3 & 0x80) == 0)
            return true;

        if (Stream.Read(buffer, 0, 1) == 0)
            throw new ExcelReaderException(Errors.ErrorBiffRecordSizeInvalid);
        byte b4 = buffer[0];
        value = ((uint)(b4 & 0x7F) << 21) | value;

        if ((b4 & 0x80) != 0)
            throw new ExcelReaderException(Errors.ErrorBiffRecordSizeInvalid);

        return true;
    }
}