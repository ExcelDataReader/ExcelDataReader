using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Excel.Log;
using ICSharpCode.SharpZipLib.Zip;
using System.Collections;

namespace Excel.Core
{
    public class MemoryWorker : IExcelWorker
    {
        #region Members and Properties

        private byte[] buffer;

        private bool disposed;

        private const string TMP = "TMP_Z";
        //private const string FOLDER_xl = "xl";
        //private const string FOLDER_worksheets = "worksheets";
        //private const string FILE_sharedStrings = "sharedStrings.{0}";
        //private const string FILE_styles = "styles.{0}";
        //private const string FILE_workbook = "workbook.{0}";
        //private const string FILE_sheet = "sheet{0}.{1}";
        //private const string FOLDER_rels = "_rels";
        //private const string FILE_rels = "workbook.{0}.rels";

        private string _exceptionMessage;

        private bool _isValid;

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

        public MemoryWorker()
        {
            MemoryData.zipfilelist = new List<List<string>>();
        }

        public bool Extract(Stream fileStream)
        {
            if (null == fileStream) return false;
            _isValid = true;
            ZipFile zipFile = null;
            try
            {
                zipFile = new ZipFile(fileStream);
                IEnumerator enumerator = zipFile.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    ZipEntry entry = (ZipEntry)enumerator.Current;
                    if (buffer == null)
                    {
                        buffer = new byte[0x1000];
                    }
                    List<string> zipfileitems = new List<string>();
                    MemoryData.excelSavedFromAccessBool = false;
                    using (StreamReader reader = new StreamReader(zipFile.GetInputStream(entry)))
                    {
                        // Create our List and Populate it
                        zipfileitems.Add(entry.ToString());
                        //if (entry.Name.Contains("thumbnail.wmf"))
                        //{
                        //    MemoryData.excelSavedFromAccessBool = true;
                        //    _isValid = false;
                        //    break;
                        //}
                        #region output for debugging
                        ////Console.Write("Name " + entry.Name.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("Type " + entry.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("AESKeySize " + entry.AESKeySize.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("CanDecompress " + entry.CanDecompress.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("CentralHeaderRequiresZip64 " + entry.CentralHeaderRequiresZip64.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("CompressedSize " + entry.CompressedSize.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("Crc " + entry.Crc.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("ExternalFileAttributes " + entry.ExternalFileAttributes.ToString());
                        ////Console.Write("\r\n");
                        //////   Console.Write("ExtraData " + entry.ExtraData.ToString());
                        ////Console.Write("Flags " + entry.Flags.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("HasCrc " + entry.HasCrc.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("IsCrypted " + entry.IsCrypted.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("IsDirectory. " + entry.IsDirectory.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("IsFile " + entry.IsFile.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("VersionMadeBy " + entry.VersionMadeBy.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("ZipFileIndex " + entry.ZipFileIndex.ToString());
                        ////Console.Write("\r\n");
                        ////Console.Write("\r\n");
                        ////Console.Write("\r\n");
                        ////Console.Write("\r\n");
                        ////Console.Write("\r\n");
                        ////Console.Write("\r\n");
                        ////Console.Write(" ----------------------------------------------------------");
                        #endregion
                        try
                        {
                            zipfileitems.Add(reader.ReadToEnd());
                        }
                        catch (Exception ex)
                        {
                            _isValid = false;
                            _exceptionMessage = ex.Message;
                            throw;
                        }
                        finally
                        {
                            MemoryData.zipfilelist.Add(zipfileitems);
                            reader.Dispose();
                        }
                    }
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
            /*try
            {
                if (!Directory.Exists(TMP)) Directory.CreateDirectory(TMP);
                foreach (List<string> file in MemoryData.zipfilelist)
                {
                    if (!String.IsNullOrEmpty(file[1]))
                    {
                        string path = Path.Combine(TMP, file[0]);
                        string folder = Path.GetDirectoryName(path);
                        if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                        File.WriteAllText(path, file[1], Encoding.UTF8);
                    }
                }
            }
            catch (Exception ex) { }*/
            return _isValid;
        }

        /// <summary>
        /// Gets the shared strings stream.
        /// </summary>
        /// <returns></returns>
        public Stream GetSharedStringsStream()
        {
            string exelcontent = "xl/sharedStrings.xml";
            return getStream(exelcontent);
        }

        /// <summary>
        /// Gets the styles stream.
        /// </summary>
        /// <returns></returns>
        public Stream GetStylesStream()
        {
            string exelcontent = "xl/styles.xml";
            return getStream(exelcontent);
        }

        /// <summary>
        /// Gets the workbook stream.
        /// </summary>
        /// <returns></returns>
        public Stream GetWorkbookStream()
        {
            string exelcontent = "xl/workbook.xml";
            return getStream(exelcontent);
        }

        /// <summary>
        /// Gets the worksheet stream.
        /// </summary>
        /// <param name="sheetId">The sheet id.</param>
        /// <returns></returns>
        public Stream GetWorksheetStream(int sheetId)
        {
            string excelcontent = "xl/worksheets/sheet" + sheetId + ".xml";
            return getStream(excelcontent);
        }

        public Stream GetWorksheetStream(string sheetPath)
        {
            //its possible sheetPath starts with /xl. in this case trim the /xl
            if (sheetPath.StartsWith("/xl/"))
                sheetPath = sheetPath.Substring(4);
            string exelcontent = "xl/" + sheetPath;
            return getStream(exelcontent);
        }

        /// <summary>
        /// Gets the workbook rels stream.
        /// </summary>
        /// <returns></returns>
        public Stream GetWorkbookRelsStream()
        {
            string exelcontent = "xl/_rels/workbook.xml.rels";
            return getStream(exelcontent);
        }

        private MemoryStream getStreamFromString(string excelstring)
        {
            MemoryStream ms = new MemoryStream();
            if (!String.IsNullOrEmpty(excelstring))
            {
                try
                {
                    Byte[] byteArray = Encoding.UTF8.GetBytes(excelstring);
                    ms = new MemoryStream(byteArray);
                    return ms;
                }
                catch (Exception ex)
                {
                    ms.Dispose();
                    _exceptionMessage = ex.Message;
                    throw;
                }
            }
            else
            {
                return null;
            }
        }

        private Stream getStream(string itemlist)
        {
            Stream data = null;
            foreach (List<string> item in MemoryData.returnlist)
            {
                if (item.Contains(itemlist))
                {
                    data = getStreamFromString(item[1]);
                    return data;
                }
            }
            return data; // nothing to return
        }

        /*private static byte[] StringToBytes(string str)
        {
            byte[] data = new byte[str.Length * 2];
            for (int i = 0; i < str.Length; ++i)
            {
                char ch = str[i]; data[i * 2] = (byte)(ch & 0xFF); data[i * 2 + 1] = (byte)((ch & 0xFF00) >> 8);
            }
            return data;
        }*/

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
                }

                buffer = null;

                disposed = true;
            }
        }

        ~MemoryWorker()
        {
            Dispose(false);
        }

        #endregion
    }

    internal class MemoryData
    {
        public static bool excelSavedFromAccessBool;

        public static List<List<String>> zipfilelist = new List<List<String>>();

        private static void zipfilelistadd(List<string> zipfileitems)
        {
            zipfilelist.Add(zipfileitems);
        }

        public static List<List<string>> returnlist
        {
            get { return zipfilelist; }
        }
    }
}
