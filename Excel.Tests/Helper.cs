using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.IO;
#if MSTEST_DEBUG || MSTEST_RELEASE
using Microsoft.VisualStudio.TestTools.UnitTesting;
#else
using NUnit.Framework;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestInitialize = NUnit.Framework.SetUpAttribute;
using TestCleanup = NUnit.Framework.TearDownAttribute;
using TestMethod = NUnit.Framework.TestAttribute;
#endif

#if LEGACY
namespace Excel.Tests
#else
namespace ExcelDataReader.Tests
#endif
{
    internal static class Helper
    {
        public static Stream GetTestWorkbook(string key)
        {
            var fileName = GetTestWorkbookPath(key);
            return new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }

        public static string GetKey(string key)
        {
            string pathFile = ConfigurationManager.AppSettings[key];
            Debug.WriteLine(pathFile);
            return pathFile;
        }

        public static double ParseDouble(string s)
        {
            return double.Parse(s, CultureInfo.InvariantCulture);
        }

        public static string GetTestWorkbookPath(string key)
        {
            string fileName = Path.Combine(GetKey("basePath"), GetKey(key));
            fileName = Path.GetFullPath(fileName);
            Assert.IsTrue(File.Exists(fileName), string.Format("By the key '{0}' the file '{1}' could not be found. Inside the Excel.Tests App.config file, edit the key basePath to be the folder where the test workbooks are located. If this is fine, check the filename that is related to the key.", key, fileName));
            return fileName;
        }

		// Merged From linked CopyStream below and Jon Skeet's ReadFully example
		public static void CopyStream(Stream input, Stream output)
		{
			byte[] buffer = new byte[16*1024];
			int read;
			while((read = input.Read (buffer, 0, buffer.Length)) > 0)
			{
				output.Write (buffer, 0, read);
			}
		}
    }
}