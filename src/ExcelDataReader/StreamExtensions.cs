#nullable enable

namespace ExcelDataReader;

internal static class StreamExtensions
{
    public static int ReadAtLeast(this Stream stream, byte[] buffer, int offset, int minimumBytes)
    {
#if NET8_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfNegative(minimumBytes);
        ArgumentOutOfRangeException.ThrowIfLessThan(buffer.Length, offset + minimumBytes);
#else
        if (minimumBytes < 0 || buffer.Length < offset + minimumBytes)
            throw new ArgumentOutOfRangeException(nameof(minimumBytes));
#endif
        int totalRead = 0;
        while (totalRead < minimumBytes)
        {
            int read = stream.Read(buffer, offset + totalRead, minimumBytes - totalRead);
            if (read == 0)
                return totalRead;

            totalRead += read;
        }

        return totalRead;
    }
}
