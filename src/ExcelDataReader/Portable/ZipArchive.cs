#if NET20
using System;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;

namespace ExcelDataReader.Core
{
    public sealed class ZipArchive : IDisposable
    {
        private readonly ZipFile _handle;

        public ZipArchive(Stream stream)
        {
            _handle = new ZipFile(stream);
        }

        public ZipEntry GetEntry(string name)
        {
            var entry = _handle.GetEntry(name);
            if (entry == null)
                return null;
            return new ZipEntry(_handle, entry);
        }

        public void Dispose()
        {
        }
    }
}

#endif
