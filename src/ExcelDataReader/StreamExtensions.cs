namespace ExcelDataReader;

internal static class StreamExtensions
{
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
