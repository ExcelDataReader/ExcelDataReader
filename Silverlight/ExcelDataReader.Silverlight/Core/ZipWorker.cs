namespace ExcelDataReader.Silverlight.Core
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using ICSharpCode.SharpZipLib.Zip;

    public class ZipWorker : IDisposable
	{
		#region Members and Properties

		private byte[] _Buffer;

		private bool _IsDisposed;

		private const string TMP = "TMP_Z";
		private const string FOLDER_xl = "xl";
		private const string FOLDER_worksheets = "worksheets";
		private const string FILE_sharedStrings = "sharedStrings.{0}";
		private const string FILE_styles = "styles.{0}";
		private const string FILE_workbook = "workbook.{0}";
		private const string FILE_sheet = "sheet{0}.{1}";
		private const string FOLDER_rels = "_rels";
		private const string FILE_rels = "workbook.{0}.rels";

		private string _exceptionMessage;
        private string _format = "xml";

		private bool _isValid;

        private Dictionary<string, MemoryStream> fakeFileSystem = new Dictionary<string, MemoryStream>();
        public Dictionary<string, MemoryStream> FakeFileSystem
        {
            get { return fakeFileSystem; }
        }


		/// <summary>
		/// Gets a value indicating whether this instance is valid.
		/// </summary>
		/// <value><c>true</c> if this instance is valid; otherwise, <c>false</c>.</value>
		public bool IsValid
		{
			get { return _isValid; }
		}

		/// <summary>
		/// Gets the exception message.
		/// </summary>
		/// <value>The exception message.</value>
		public string ExceptionMessage
		{
			get { return _exceptionMessage; }
		}

		#endregion

		public ZipWorker(){ }

		/// <summary>
		/// Extracts the specified zip file stream.
		/// </summary>
		/// <param name="fileStream">The zip file stream.</param>
		/// <returns></returns>
		public void Extract(Stream fileStream)
		{
			_isValid = true;

			ZipFile zipFile = null;

			try
			{
				zipFile = new ZipFile(fileStream);

				IEnumerator enumerator = zipFile.GetEnumerator();

				while (enumerator.MoveNext())
				{
					ZipEntry entry = (ZipEntry)enumerator.Current;

					ExtractZipEntry(zipFile, entry);
				}
			}
			catch (Exception ex)
			{
				_isValid = false;
				_exceptionMessage = ex.Message;
			}
			finally
			{
				fileStream.Close();

				if (null != zipFile) zipFile.Close();
			}
		}

		/// <summary>
		/// Gets the shared strings stream.
		/// </summary>
		/// <returns></returns>
        public byte[] GetSharedStringsByteArray()
		{
            return GetByteArray(FOLDER_xl + "/" + string.Format(FILE_sharedStrings, _format));
		}

		/// <summary>
		/// Gets the styles stream.
		/// </summary>
		/// <returns></returns>
        public byte[] GetStylesByteArray()
		{
            return GetByteArray(FOLDER_xl + "/" + string.Format(FILE_styles, _format));
		}

		/// <summary>
		/// Gets the workbook stream.
		/// </summary>
		/// <returns></returns>
		public byte[] GetWorkbookByteArray()
		{
            return GetByteArray(FOLDER_xl + "/" + string.Format(FILE_workbook, _format));
		}

		/// <summary>
		/// Gets the worksheet stream.
		/// </summary>
		/// <param name="sheetId">The sheet id.</param>
		/// <returns></returns>
        public byte[] GetWorksheetByteArray(int sheetId)
		{
            return GetByteArray(FOLDER_xl + "/" + FOLDER_worksheets + "/" + string.Format(FILE_sheet, sheetId, _format));
		}

        public byte[] GetWorksheetByteArray(string sheetPath)
        {
            return GetByteArray(FOLDER_xl + "/" + sheetPath);
        }


		/// <summary>
		/// Gets the workbook rels stream.
		/// </summary>
		/// <returns></returns>
        public byte[] GetWorkbookRelsByteArray()
		{
            return GetByteArray(FOLDER_xl + "/" + FOLDER_rels + "/" + string.Format(FILE_rels, _format));
		}

		private void ExtractZipEntry(ZipFile zipFile, ZipEntry entry)
		{
			if (!entry.IsCompressionMethodSupported() || string.IsNullOrEmpty(entry.Name) || !entry.IsFile) return;

            using (MemoryStream stream = new MemoryStream())
            {
                if (_Buffer == null)
                {
                    _Buffer = new byte[0x1000];
                }

                Stream inputStream = zipFile.GetInputStream(entry);

                int count;
                while ((count = inputStream.Read(_Buffer, 0, _Buffer.Length)) > 0)
                {
                    stream.Write(_Buffer, 0, count);
                }

                stream.Flush();
                stream.Close();

                fakeFileSystem.Add(entry.Name, stream);
            }
		}

		private byte[] GetByteArray(string filePath)
		{
            if (fakeFileSystem.ContainsKey(filePath))
            {
                return fakeFileSystem[filePath].ToArray();
            }
            else
            {
                return null;
            }
		}

		#region IDisposable Members

		public void Dispose()
		{
			Dispose(true);

			GC.SuppressFinalize(this);
		}

		private void Dispose(bool disposing)
		{
			// Check to see if Dispose has already been called.
			if (!this._IsDisposed)
			{
				_Buffer = null;

				_IsDisposed = true;
			}
		}

		~ZipWorker()
		{
			Dispose(false);
		}

		#endregion
	}
}