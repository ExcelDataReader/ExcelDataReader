using System;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using System.Collections;
using ExcelDataReader.Portable.Async;
using ExcelDataReader.Portable.IO;
using ExcelDataReader.Portable.IO.PCLStorage;
using ExcelDataReader.Portable.Log;
using PCLStorage;

namespace ExcelDataReader.Portable.Core
{
	public class ZipWorker : IExcelWorker
	{
	    private readonly IFileSystem fileSystem;
	    private readonly IFileHelper fileHelper;

	    #region Members and Properties

		private byte[] buffer;

		private bool disposed;
		private bool isCleaned;

		private const string TMP = "TMP_Z";
		private const string FOLDER_xl = "xl";
		private const string FOLDER_worksheets = "worksheets";
		private const string FILE_sharedStrings = "sharedStrings.{0}";
		private const string FILE_styles = "styles.{0}";
		private const string FILE_workbook = "workbook.{0}";
		private const string FILE_sheet = "sheet{0}.{1}";
		private const string FOLDER_rels = "_rels";
		private const string FILE_rels = "workbook.{0}.rels";

		private string tempPath;
		private string exceptionMessage;
		private string xlPath;
		private string format = "xml";

		private bool isValid;
	    private string folderName;
	    private IFolder rootFolder;
	    //private bool _isBinary12Format;

		/// <summary>
		/// Gets a value indicating whether this instance is valid.
		/// </summary>
		/// <value><c>true</c> if this instance is valid; otherwise, <c>false</c>.</value>
		public bool IsValid
		{
			get { return isValid; }
		}

		/// <summary>
		/// Gets the temp path for extracted files.
		/// </summary>
		/// <value>The temp path for extracted files.</value>
		public string TempPath
		{
			get { return tempPath; }
		}

		/// <summary>
		/// Gets the exception message.
		/// </summary>
		/// <value>The exception message.</value>
		public string ExceptionMessage
		{
			get { return exceptionMessage; }
		}

		#endregion

		public ZipWorker(IFileSystem fileSystem, IFileHelper fileHelper)
		{
		    this.fileSystem = fileSystem;
		    this.fileHelper = fileHelper;
		}

	    /// <summary>
		/// Extracts the specified zip file stream.
		/// </summary>
		/// <param name="fileStream">The zip file stream.</param>
		/// <returns></returns>
		public async Task<bool> Extract(Stream fileStream)
		{
			if (null == fileStream) return false;

			await CleanFromTempAsync(false);

			await NewTempPath();

			isValid = true;

            ZipArchive zipFile = null;

			try
			{
			    zipFile = new ZipArchive(fileStream);

			    IEnumerator enumerator = zipFile.Entries.GetEnumerator();

			    while (enumerator.MoveNext())
			    {
			        var entry = (ZipArchiveEntry) enumerator.Current;

			        await ExtractZipEntry(zipFile, entry);
			    }
			}
			catch (InvalidDataException ex)
			{
                isValid = false;
                exceptionMessage = ex.Message;

                CleanFromTemp(true);
			}
			catch (Exception ex)
			{
				CleanFromTemp(true); //true tells CleanFromTemp not to raise an IO Exception if this operation fails. If it did then the real error here would be masked
			    throw;
			}
			finally
			{
				fileStream.Dispose();

                if (null != zipFile) zipFile.Dispose();
			}

			return isValid && await CheckFolderTree();
		}

		/// <summary>
		/// Gets the shared strings stream.
		/// </summary>
		/// <returns></returns>
        public async Task<Stream> GetSharedStringsStream()
		{
			return await GetStream(Path.Combine(xlPath, string.Format(FILE_sharedStrings, format)));
		}

		/// <summary>
		/// Gets the styles stream.
		/// </summary>
		/// <returns></returns>
        public async Task<Stream> GetStylesStream()
		{
            return await GetStream(Path.Combine(xlPath, string.Format(FILE_styles, format)));
		}

		/// <summary>
		/// Gets the workbook stream.
		/// </summary>
		/// <returns></returns>
		public async Task<Stream> GetWorkbookStream()
		{
            return await GetStream(Path.Combine(xlPath, string.Format(FILE_workbook, format)));
		}

		/// <summary>
		/// Gets the worksheet stream.
		/// </summary>
		/// <param name="sheetId">The sheet id.</param>
		/// <returns></returns>
		public async Task<Stream> GetWorksheetStream(int sheetId)
		{
			return await GetStream(Path.Combine(
				Path.Combine(xlPath, FOLDER_worksheets),
				string.Format(FILE_sheet, sheetId, format)));
		}

        public async Task<Stream> GetWorksheetStream(string sheetPath)
        {
			//its possible sheetPath starts with /xl. in this case trim the /xl
	        if (sheetPath.StartsWith("/xl/"))
		        sheetPath = sheetPath.Substring(4);
            return await GetStream(Path.Combine(xlPath, sheetPath));
        }


		/// <summary>
		/// Gets the workbook rels stream.
		/// </summary>
		/// <returns></returns>
		public async Task<Stream> GetWorkbookRelsStream()
		{
			return await GetStream(Path.Combine(xlPath, Path.Combine(FOLDER_rels, string.Format(FILE_rels, format))));
		}

		private async Task CleanFromTempAsync(bool catchIoError)
		{
			if (string.IsNullOrEmpty(tempPath)) return;

			isCleaned = true;

			try
			{
			    var exists = await fileSystem.LocalStorage.CheckExistsAsync(tempPath);
                if (exists == ExistenceCheckResult.FolderExists)
				{
				    var dir = await fileSystem.GetFolderFromPathAsync(tempPath);
                    await dir.DeleteFolderAndContentsAsync();
				}
			}
			catch (IOException ex)
			{
			    this.Log().Error(ex.Message);
				if (!catchIoError)
					throw;
			}
			
		}

        private void CleanFromTemp(bool catchIoError)
        {
            //todo: not sure about this because CleanFromTemp can get called in Exception handling and dispose
            //I think it's ok because we wait for it here
            AsyncHelper.RunSync(() => CleanFromTempAsync(catchIoError));

        }


        private async Task ExtractZipEntry(ZipArchive zipFile, ZipArchiveEntry entry)
		{
			if (string.IsNullOrEmpty(entry.Name)) return;

            

			string tPath = Path.Combine(tempPath, entry.Name);
            
            //string path = entry..IsDirectory ? tPath : Path.GetDirectoryName(Path.GetFullPath(tPath));

            var containingDirectoryName = Path.GetDirectoryName(entry.FullName);
            IFolder containingFolder = null;
 
            //get or create containing directory
            if (string.IsNullOrEmpty(containingDirectoryName))
            {
                containingFolder = rootFolder;
            }
            else
            {
                //this is a sub folder so make sure it is created
                var folderExists = await rootFolder.CheckExistsAsync(containingDirectoryName);

                if (folderExists == ExistenceCheckResult.NotFound)
                {
                    await rootFolder.CreateFolderAsync(containingDirectoryName, CreationCollisionOption.ReplaceExisting);
                }

                //get reference to the folder
                containingFolder = await rootFolder.GetFolderAsync(containingDirectoryName);
            }

            //create the file
            var fileExists = await containingFolder.CheckExistsAsync(entry.Name);

            if (fileExists == ExistenceCheckResult.NotFound)
            {
                await containingFolder.CreateFileAsync(entry.Name, CreationCollisionOption.ReplaceExisting);
            }

            var file = await containingFolder.GetFileAsync(entry.Name);

			using (var stream = await file.OpenAsync(FileAccess.ReadAndWrite))
			{
				if (buffer == null)
				{
					buffer = new byte[0x1000];
				}

				using(var inputStream = entry.Open())
				{
					int count;
					while ((count = inputStream.Read(buffer, 0, buffer.Length)) > 0)
					{
						stream.Write(buffer, 0, count);
					}
				}

					

				stream.Flush();
			}
		}

		private async Task NewTempPath()
		{
		    var tempID = Guid.NewGuid().ToString("N");
		    folderName = TMP + DateTime.Now.ToFileTimeUtc().ToString() + tempID;
            tempPath = Path.Combine(fileHelper.GetTempPath(), folderName);

            //ensure root folder created
		    var rootExists = await fileSystem.LocalStorage.CheckExistsAsync(tempPath);
            if (rootExists == ExistenceCheckResult.NotFound)
            {
                await fileSystem.LocalStorage.CreateFolderAsync(tempPath, CreationCollisionOption.ReplaceExisting);
            }
		    rootFolder = await fileSystem.GetFolderFromPathAsync(tempPath);
			isCleaned = false;

            this.Log().Debug("Using temp path {0}", tempPath);

		}

		private async Task<bool> CheckFolderTree()
		{
			xlPath = Path.Combine(tempPath, FOLDER_xl);

            var existsXlPath = await fileSystem.LocalStorage.CheckExistsAsync(xlPath) == ExistenceCheckResult.FolderExists;
            var existsWorksheetPath = await fileSystem.LocalStorage.CheckExistsAsync(Path.Combine(xlPath, FOLDER_worksheets)) == ExistenceCheckResult.FolderExists;
            var existsWorkbook = await fileSystem.LocalStorage.CheckExistsAsync(Path.Combine(xlPath, FILE_workbook)) == ExistenceCheckResult.FileExists;
            var existsStyles = await fileSystem.LocalStorage.CheckExistsAsync(Path.Combine(xlPath, FILE_styles)) == ExistenceCheckResult.FileExists;

            return existsXlPath &&
                existsWorksheetPath &&
                existsWorkbook &&
                existsStyles;
		}

		private async Task<Stream> GetStream(string filePath)
		{

		    var fileExists = await fileSystem.LocalStorage.CheckExistsAsync(filePath) == ExistenceCheckResult.FileExists;
            if (fileExists)
            {
                var file = await fileSystem.GetFileFromPathAsync(filePath);
                return await file.OpenAsync(FileAccess.Read);
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
			if (!this.disposed)
			{
				if (disposing)
				{
					if (!isCleaned)
						CleanFromTemp(false);
				}

				buffer = null;

				disposed = true;
			}
		}

		~ZipWorker()
		{
			Dispose(false);
		}

		#endregion
	}
}