using System.IO;
using ExcelDataReader.Core.BinaryFormat;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader
{
    /// <summary>
    /// ExcelDataReader Class
    /// </summary>
    internal partial class ExcelBinaryReader : ExcelDataReader<XlsWorkbook, XlsWorksheet>
    {
        private const string DirectoryEntryWorkbook = "Workbook";
        private const string DirectoryEntryBook = "Book";

        public ExcelBinaryReader(Stream stream, ExcelReaderConfiguration configuration)
            : base(configuration)
        {
            Stream = stream;
            Document = new XlsDocument(stream);
            Workbook = ReadWorkbook();

            // By default, the data reader is positioned on the first result.
            Reset();
        }

        private Stream Stream { get; set; }

        private XlsDocument Document { get; set; }

        public override void Close()
        {
            base.Close();

            Stream?.Dispose();
            Stream = null;
            Workbook = null;
            Document = null;
        }

        private XlsWorkbook ReadWorkbook()
        {
            XlsDirectoryEntry workbookEntry = Document.FindEntry(DirectoryEntryWorkbook) ?? Document.FindEntry(DirectoryEntryBook);

            if (workbookEntry == null)
            {
                throw new ExcelReaderException(Errors.ErrorStreamWorkbookNotFound);
            }

            if (workbookEntry.EntryType != STGTY.STGTY_STREAM)
            {
                throw new ExcelReaderException(Errors.ErrorWorkbookIsNotStream);
            }

            var bytes = Document.ReadStream(Stream, workbookEntry.StreamFirstSector, (int)workbookEntry.StreamSize, workbookEntry.IsEntryMiniStream);

            return new XlsWorkbook(bytes, Configuration.ConvertOaDate, Configuration.ReadOption);
        }
    }
}
