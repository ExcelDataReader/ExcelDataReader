#if NET20
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;

namespace ExcelDataReader.Core
{
    internal sealed class ZipArchive : IDisposable
    {
        private readonly ZipFile _handle;

        public ZipArchive(Stream stream)
        {
            if (stream.CanSeek)
            {
                _handle = new ZipFile(stream);
            }
            else
            {
                // Password protected xlsx using "Standard Encryption" come as a non-seekable CryptoStream.
                // Must wrap in a MemoryStream to load
                var memoryStream = ReadToMemoryStream(stream);
                _handle = new ZipFile(memoryStream);
            }

            var entries = new List<ZipArchiveEntry>();
            foreach (ZipEntry entry in _handle)
                entries.Add(new ZipArchiveEntry(_handle, entry));
            Entries = new ReadOnlyCollection<ZipArchiveEntry>(entries);
        }

        public ReadOnlyCollection<ZipArchiveEntry> Entries { get; }

        public void Dispose()
        {
            (_handle as IDisposable)?.Dispose();
        }

        private static MemoryStream ReadToMemoryStream(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            int read;
            var ms = new MemoryStream();
            while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                ms.Write(buffer, 0, read);
            }

            ms.Position = 0;
            return ms;
        }
    }
}
#endif
