using System.Diagnostics;
using System.Globalization;
using System.IO;
using NUnit.Framework;

namespace ExcelDataReader.Tests
{
    internal static class Helper {

		public static Stream GetTestWorkbook(string key) {
            var fileName = GetTestWorkbookPath(key);
            return new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }

        public static string GetKey(string key)
        {
            string pathFile = Configuration.AppSettings[key];
            Debug.WriteLine(pathFile);
            return pathFile;
        }

		public static string GetTestWorkbookPath(string key)
        {
			var fileName = Path.Combine("../Resources", GetKey(key));
            //string fileName = Path.Combine(GetKey("basePath"), GetKey(key));
            fileName = Path.GetFullPath(fileName);
            Assert.IsTrue(File.Exists(fileName), string.Format("By the key '{0}' the file '{1}' could not be found. Inside the Excel.Tests App.config file, edit the key basePath to be the folder where the test workbooks are located. If this is fine, check the filename that is related to the key.", key, fileName));
            return fileName;
        }

    }
}