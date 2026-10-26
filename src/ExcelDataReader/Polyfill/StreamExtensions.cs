namespace System.IO;

internal static class StreamExtensions
{
#if NETSTANDARD2_1_OR_GREATER && !NET8_0_OR_GREATER
    public static int ReadAtLeast(this Stream stream, Span<byte> buffer, int minimumBytes, bool throwOnEndOfStream = true)
    {
        if (minimumBytes < 0 || minimumBytes > buffer.Length)
            throw new ArgumentOutOfRangeException(nameof(minimumBytes));

        int totalRead = 0;
        while (totalRead < minimumBytes)
        {
            int read = stream.Read(buffer[totalRead..]);
            if (read == 0)
            {
                if (throwOnEndOfStream)
                    throw new EndOfStreamException();
                break;
            }

            totalRead += read;
        }

        return totalRead;
    }
#endif

    public static int ReadAtLeast(this Stream stream, byte[] buffer, int offset, int count, int minimumBytes, bool throwOnEndOfStream = true)
    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        return stream.ReadAtLeast(buffer.AsSpan(offset, count), minimumBytes, throwOnEndOfStream);
#else
        ArgumentNullException.ThrowIfNull(buffer);
        if (offset < 0 || offset > buffer.Length)
            throw new ArgumentOutOfRangeException(nameof(offset));
        if (count < 0 || count > buffer.Length - offset)
            throw new ArgumentOutOfRangeException(nameof(count));
        if (minimumBytes < 0 || minimumBytes > count)
            throw new ArgumentOutOfRangeException(nameof(minimumBytes));

        int totalRead = 0;
        while (totalRead < minimumBytes)
        {
            int read = stream.Read(buffer, offset + totalRead, count - totalRead);
            if (read == 0)
            {
                if (throwOnEndOfStream)
                    throw new EndOfStreamException();
                break;
            }

            totalRead += read;
        }

        return totalRead;
#endif
    }
}