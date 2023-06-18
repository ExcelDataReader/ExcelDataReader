using System.Collections.Generic;
using System.IO;
#if MSTEST_DEBUG || MSTEST_RELEASE
using Microsoft.VisualStudio.TestTools.UnitTesting;
#else
using NUnit.Framework;
#endif

namespace ExcelDataReader.Tests
{
    internal static class Configuration
    {
        public static Stream GetTestWorkbook(string key)
        {
            var fileName = GetTestWorkbookPath(key);
            return new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }

        public static string GetTestWorkbookPath(string key)
        {
            var resources = Path.Combine(TestContext.CurrentContext.TestDirectory, "../../../../Resources");
            var path = Path.Combine(resources, key);
            path = Path.GetFullPath(path);
            Assert.IsTrue(File.Exists(path), $"File not found: '{path}'.");
            return path;
        }

        public static ExcelDataSetConfiguration NoColumnNamesConfiguration = new()
        {
            ConfigureDataTable = reader => new()
            {
                UseHeaderRow = false
            }
        };

        public static ExcelDataSetConfiguration FirstRowColumnNamesConfiguration = new()
        {
            ConfigureDataTable = reader => new()
            {
                UseHeaderRow = true
            }
        };

        public static ExcelDataSetConfiguration FirstRowColumnNamesPrefixConfiguration = new()
        {
            ConfigureDataTable = reader => new()
            {
                UseHeaderRow = true,
                EmptyColumnNamePrefix = "Prefix"
            }
        };
    }
}
