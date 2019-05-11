using System;
using System.IO;
#if !NET20
using System.IO.Compression;
#endif

namespace ExcelDataReader.Core
{
    internal partial class ZipWorker : IDisposable
    {
        public ZipArchive MyZipWorker(Stream stream)
        {
            return new ZipArchive(stream);
        }
    }
}
