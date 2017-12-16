using System;
using System.IO;
#if !NET20
using System.IO.Compression;
#endif

namespace ExcelDataReader.Core
{
    internal partial class ZipWorker : IDisposable
    {
        private const string FileSharedStrings = "xl/sharedStrings.{0}";
        private const string FileStyles = "xl/styles.{0}";
        private const string FileWorkbook = "xl/workbook.{0}";
        private const string FileSheet = "xl/worksheets/sheet{0}.{1}";
        private const string FileRels = "xl/_rels/workbook.{0}.rels";

        private const string Format = "xml";

        private bool _disposed;
        private Stream _zipStream;
        private ZipArchive _zipFile;

        /// <summary>
        /// Initializes a new instance of the <see cref="ZipWorker"/> class. 
        /// </summary>
        /// <param name="fileStream">The zip file stream.</param>
        public ZipWorker(Stream fileStream)
        {
            _zipStream = fileStream ?? throw new ArgumentNullException(nameof(fileStream));
            _zipFile = new ZipArchive(fileStream);
        }

        /// <summary>
        /// Gets the shared strings stream.
        /// </summary>
        /// <returns>The shared strings stream.</returns>
        public Stream GetSharedStringsStream()
        {
            var zipEntry = FindEntry(string.Format(FileSharedStrings, Format));
            return zipEntry?.Open();
        }

        /// <summary>
        /// Gets the styles stream.
        /// </summary>
        /// <returns>The styles stream.</returns>
        public Stream GetStylesStream()
        {
            var zipEntry = FindEntry(string.Format(FileStyles, Format));
            return zipEntry?.Open();
        }

        /// <summary>
        /// Gets the workbook stream.
        /// </summary>
        /// <returns>The workbook stream.</returns>
        public Stream GetWorkbookStream()
        {
            var zipEntry = FindEntry(string.Format(FileWorkbook, Format));
            return zipEntry?.Open();
        }

        /// <summary>
        /// Gets the worksheet stream.
        /// </summary>
        /// <param name="sheetId">The sheet id.</param>
        /// <returns>The worksheet stream.</returns>
        public Stream GetWorksheetStream(int sheetId)
        {
            var zipEntry = FindEntry(string.Format(FileSheet, sheetId, Format));
            return zipEntry?.Open();
        }

        public Stream GetWorksheetStream(string sheetPath)
        {
            // its possible sheetPath starts with /xl. in this case trim the /
            // see the test "Issue_11522_OpenXml"
            if (sheetPath.StartsWith("/xl/", StringComparison.OrdinalIgnoreCase))
                sheetPath = sheetPath.Substring(1);
            else
                sheetPath = "xl/" + sheetPath;

            var zipEntry = FindEntry(sheetPath);
            return zipEntry?.Open();
        }

        /// <summary>
        /// Gets the workbook rels stream.
        /// </summary>
        /// <returns>The rels stream.</returns>
        public Stream GetWorkbookRelsStream()
        {
            var zipEntry = FindEntry(string.Format(FileRels, Format));
            return zipEntry?.Open();
        }

        private ZipArchiveEntry FindEntry(string name)
        {
            foreach (var entry in _zipFile.Entries)
            {
                if (entry.FullName.Equals(name, StringComparison.OrdinalIgnoreCase))
                    return entry;
            }

            return null;
        }
    }

    internal partial class ZipWorker
    {
        ~ZipWorker()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!_disposed)
            {
                if (disposing)
                {
                    if (_zipFile != null)
                    {
                        _zipFile.Dispose();
                        _zipFile = null;
                    }

                    if (_zipStream != null)
                    {
                        _zipStream.Dispose();
                        _zipStream = null;
                    }
                }

                _disposed = true;
            }
        }
    }
}