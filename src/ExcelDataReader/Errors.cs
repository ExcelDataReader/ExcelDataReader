namespace ExcelDataReader
{
    internal static class Errors
    {
        public const string ErrorStreamWorkbookNotFound = "Neither stream 'Workbook' nor 'Book' was found in file.";
        public const string ErrorWorkbookIsNotStream = "Workbook directory entry is not a Stream.";
        public const string ErrorWorkbookGlobalsInvalidData = "Error reading Workbook Globals - Stream has invalid data.";
        public const string ErrorFatBadSector = "Error reading as FAT table : There's no such sector in FAT.";
        public const string ErrorFatRead = "Error reading stream from FAT area.";
        public const string ErrorHeaderSignature = "Invalid file signature.";
        public const string ErrorHeaderOrder = "Invalid byte order specified in header.";
        public const string ErrorBiffRecordSize = "Buffer size is less than minimum BIFF record size.";
        public const string ErrorBiffBufferSize = "BIFF Stream error: Buffer size is less than entry length.";
        public const string ErrorBiffIlegalBefore = "BIFF Stream error: Moving before stream start.";
        public const string ErrorBiffIlegalAfter = "BIFF Stream error: Moving after stream end.";

        public const string ErrorDirectoryEntryArray = "Directory Entry error: Array is too small.";
    }
}
