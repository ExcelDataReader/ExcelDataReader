#if NET20
using System.IO;

using ICSharpCode.SharpZipLib.Zip;

namespace ExcelDataReader.Core
{
    internal sealed class ZipEntry
    {
        private readonly ZipFile _handle;
        private readonly ICSharpCode.SharpZipLib.Zip.ZipEntry _entry;

        internal ZipEntry(ZipFile handle, ICSharpCode.SharpZipLib.Zip.ZipEntry entry)
        {
            _handle = handle;
            _entry = entry;
        }

        public Stream Open()
        {
            return _handle.GetInputStream(_entry);
        }
    }
}
#endif