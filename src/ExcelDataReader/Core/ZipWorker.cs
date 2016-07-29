using System;
using System.IO;
#if !NET20
using System.IO.Compression;
#endif

namespace ExcelDataReader.Core
{
	public class ZipWorker : IDisposable
	{
#region Members and Properties

		private bool disposed;
		private const string FILE_sharedStrings = "xl/sharedStrings.{0}";
		private const string FILE_styles = "xl/styles.{0}";
		private const string FILE_workbook = "xl/workbook.{0}";
		private const string FILE_sheet = "xl/worksheets/sheet{0}.{1}";
		private const string FILE_rels = "xl/_rels/workbook.{0}.rels";
		private string _format = "xml";
		private Stream zipStream;
		private ZipArchive zipFile;

#endregion

		public ZipWorker() {
		}

		/// <summary>
		/// Extracts the specified zip file stream.
		/// </summary>
		/// <param name="fileStream">The zip file stream.</param>
		/// <returns></returns>
		public bool Open(Stream fileStream) {
			if (null == fileStream)
				return false;
			zipStream = fileStream;
			zipFile = new ZipArchive(fileStream);
			return true;
		}

		/// <summary>
		/// Gets the shared strings stream.
		/// </summary>
		/// <returns></returns>
		public Stream GetSharedStringsStream() {
			var zipEntry = zipFile.GetEntry(string.Format(FILE_sharedStrings, _format));
			return zipEntry != null ? zipEntry.Open() : null;
		}

		/// <summary>
		/// Gets the styles stream.
		/// </summary>
		/// <returns></returns>
		public Stream GetStylesStream() {
			var zipEntry = zipFile.GetEntry(string.Format(FILE_styles, _format));
			return zipEntry != null ? zipEntry.Open() : null;
		}

		/// <summary>
		/// Gets the workbook stream.
		/// </summary>
		/// <returns></returns>
		public Stream GetWorkbookStream() {
			var zipEntry = zipFile.GetEntry(string.Format(FILE_workbook, _format));
			return zipEntry.Open();
		}

		/// <summary>
		/// Gets the worksheet stream.
		/// </summary>
		/// <param name="sheetId">The sheet id.</param>
		/// <returns></returns>
		public Stream GetWorksheetStream(int sheetId) {
			var zipEntry = zipFile.GetEntry(string.Format(FILE_sheet, sheetId, _format));
			return zipEntry.Open();
		}

		public Stream GetWorksheetStream(string sheetPath) {
			// its possible sheetPath starts with /xl. in this case trim the /
			// see the test "Issue_11522_OpenXml"
			if (sheetPath.StartsWith("/xl/"))
				sheetPath = sheetPath.Substring(1);
			else
				sheetPath = "xl/" + sheetPath;

			var zipEntry = zipFile.GetEntry(sheetPath);
			return zipEntry.Open();
		}


		/// <summary>
		/// Gets the workbook rels stream.
		/// </summary>
		/// <returns></returns>
		public Stream GetWorkbookRelsStream() {
			var zipEntry = zipFile.GetEntry(string.Format(FILE_rels, _format));
			return zipEntry.Open();
		}

#region IDisposable Members

		public void Dispose() {
			Dispose(true);

			GC.SuppressFinalize(this);
		}

		private void Dispose(bool disposing) {
			// Check to see if Dispose has already been called.
			if (!this.disposed) {
				if (disposing) {
					if (zipFile != null) {
						zipFile.Dispose();
						zipFile = null;
					}
					if (zipStream != null) {
						zipStream.Dispose();
						zipStream = null;
					}
				}

				disposed = true;
			}
		}

		~ZipWorker() {
			Dispose(false);
		}

#endregion
	}
}