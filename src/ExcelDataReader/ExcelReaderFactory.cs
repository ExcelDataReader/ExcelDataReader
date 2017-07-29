using System;
using System.IO;
using ExcelDataReader.Core.BinaryFormat;
using ExcelDataReader.Core.CompoundFormat;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader
{
    /// <summary>
    /// The ExcelReader Factory
    /// </summary>
    public static class ExcelReaderFactory
    {
        private const string DirectoryEntryWorkbook = "Workbook";
        private const string DirectoryEntryBook = "Book";
        private const string DirectoryEntryEncryptedPackage = "EncryptedPackage";
        private const string DirectoryEntryEncryptionInfo = "EncryptionInfo";

        /// <summary>
        /// Creates an instance of <see cref="ExcelBinaryReader"/> or <see cref="ExcelOpenXmlReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="configuration">The configuration object.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateReader(Stream fileStream, ExcelReaderConfiguration configuration = null)
        { 
            var probe = new byte[8];
            fileStream.Read(probe, 0, probe.Length);
            fileStream.Seek(0, SeekOrigin.Begin);

            if (CompoundDocument.IsCompoundDocument(probe))
            {
                // Can be BIFF5-8 or password protected OpenXml
                var document = new CompoundDocument(fileStream);
                if (TryGetWorkbook(fileStream, document, out var bytes))
                {
                    return new ExcelBinaryReader(bytes, configuration);
                }
                else if (TryGetEncryptedPackage(fileStream, document, "password", out var stream))
                {
                    return new ExcelOpenXmlReader(stream, configuration);
                }
            }

            if (XlsWorkbook.IsRawBiffStream(probe))
            {
                var bytes = new byte[fileStream.Length];
                fileStream.Read(bytes, 0, (int)fileStream.Length);
                return new ExcelBinaryReader(bytes, configuration);
            }

            // zip files start with 'PK'
            if (probe[0] == 0x50 && probe[1] == 0x4B)
            {
                return CreateOpenXmlReader(fileStream, configuration);
            }

            throw new NotSupportedException("Unknown file format");
        }

        /// <summary>
        /// Creates an instance of <see cref="ExcelBinaryReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="configuration">The configuration object.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateBinaryReader(Stream fileStream, ExcelReaderConfiguration configuration = null)
        {
            var probe = new byte[8];
            fileStream.Read(probe, 0, probe.Length);
            fileStream.Seek(0, SeekOrigin.Begin);

            byte[] bytes;
            if (CompoundDocument.IsCompoundDocument(probe))
            {
                var document = new CompoundDocument(fileStream);
                if (TryGetWorkbook(fileStream, document, out bytes))
                {
                    return new ExcelBinaryReader(bytes, configuration);
                }
            }

            bytes = new byte[fileStream.Length];
            fileStream.Read(bytes, 0, (int)fileStream.Length);
            return new ExcelBinaryReader(bytes, configuration);
        }

        /// <summary>
        /// Creates an instance of <see cref="ExcelOpenXmlReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="configuration">The reader configuration -or- <see langword="null"/> to use the default configuration.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateOpenXmlReader(Stream fileStream, ExcelReaderConfiguration configuration = null)
        {
            return new ExcelOpenXmlReader(fileStream, configuration);
        }

        private static bool TryGetWorkbook(Stream fileStream, CompoundDocument document, out byte[] bytes)
        {
            var workbookEntry = document.FindEntry(DirectoryEntryWorkbook) ?? document.FindEntry(DirectoryEntryBook);
            if (workbookEntry != null)
            {
                if (workbookEntry.EntryType != STGTY.STGTY_STREAM)
                {
                    throw new ExcelReaderException(Errors.ErrorWorkbookIsNotStream);
                }

                bytes = document.ReadStream(fileStream, workbookEntry.StreamFirstSector, (int)workbookEntry.StreamSize, workbookEntry.IsEntryMiniStream);
                return true;
            }

            bytes = null;
            return false;
        }

        private static bool TryGetEncryptedPackage(Stream fileStream, CompoundDocument document, string password, out Stream stream)
        {
            var encryptedPackage = document.FindEntry(DirectoryEntryEncryptedPackage);
            var encryptionInfo = document.FindEntry(DirectoryEntryEncryptionInfo);

            if (encryptedPackage == null || encryptionInfo == null)
            {
                stream = null;
                return false;
            }

            /*
            var packageBytes = document.ReadStream(fileStream, encryptedPackage.StreamFirstSector, (int)encryptedPackage.StreamSize, encryptedPackage.IsEntryMiniStream);
            var infoBytes = document.ReadStream(fileStream, encryptionInfo.StreamFirstSector, (int)encryptionInfo.StreamSize, encryptionInfo.IsEntryMiniStream);

            TODO ...

            */
            stream = null;
            return false;
        }
    }
}
