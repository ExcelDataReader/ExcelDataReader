namespace Excel
{
	internal static class Errors
	{
		public const string ErrorStreamWorkbookNotFound = "Error: Neither stream 'Workbook' nor 'Book' was found in file.";
		public const string ErrorWorkbookIsNotStream = "Error: Workbook directory entry is not a Stream.";
		public const string ErrorWorkbookGlobalsInvalidData = "Error reading Workbook Globals - Stream has invalid data.";
		public const string ErrorFATBadSector = "Error reading as FAT table : There's no such sector in FAT.";
		public const string ErrorFATRead = "Error reading stream from FAT area.";
		public const string ErrorHeaderSignature = "Error: Invalid file signature.";
		public const string ErrorHeaderOrder = "Error: Invalid byte order specified in header.";
		public const string ErrorBIFFRecordSize = "Error: Buffer size is less than minimum BIFF record size.";
		public const string ErrorBIFFBufferSize = "BIFF Stream error: Buffer size is less than entry length.";
		public const string ErrorBIFFIlegalBefore = "BIFF Stream error: Moving before stream start.";
		public const string ErrorBIFFIlegalAfter = "BIFF Stream error: Moving after stream end.";

		public const string ErrorDirectoryEntryArray = "Directory Entry error: Array is too small.";
	}
}
