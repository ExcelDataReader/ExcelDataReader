using System;
using System.IO;

namespace ExcelDataReader
{
    internal static class StreamExtensions
    {
        public static int ReadAtLeast(this Stream stream, byte[] buffer, int offset, int minimumBytes)
        {
            if (minimumBytes < 0)
                throw new ArgumentOutOfRangeException(nameof(minimumBytes));
            if (buffer.Length < offset + minimumBytes)
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
        }
    }
}
