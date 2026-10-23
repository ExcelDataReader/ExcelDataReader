namespace ExcelDataReader;

internal static class StreamExtensions
{
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
    public static int ReadAtLeast(this Stream stream, Span<byte> buffer)
    {
#if NET8_0_OR_GREATER
        return stream.ReadAtLeast(buffer, buffer.Length, throwOnEndOfStream: false);
#else
        int totalRead = 0;
        while (totalRead < buffer.Length)
        {
            int read = stream.Read(buffer[totalRead..]);
            if (read == 0)
                break;

            totalRead += read;
        }

        return totalRead;
#endif
    }
#endif

    public static int ReadAtLeast(this Stream stream, byte[] buffer, int offset, int minimumBytes)
    {
#if NET8_0_OR_GREATER
        // Delegate to the BCL overload (Stream.ReadAtLeast added in .NET 7).
        return stream.ReadAtLeast(buffer.AsSpan(offset, minimumBytes), minimumBytes, throwOnEndOfStream: false);
#else
        if (minimumBytes < 0 || buffer.Length < offset + minimumBytes)
            throw new ArgumentOutOfRangeException(nameof(minimumBytes));

        int totalRead = 0;
        while (totalRead < minimumBytes)
        {
            int read = stream.Read(buffer, offset + totalRead, minimumBytes - totalRead);
            if (read == 0)
                return totalRead;

            totalRead += read;
        }

        return totalRead;
#endif
    }
}