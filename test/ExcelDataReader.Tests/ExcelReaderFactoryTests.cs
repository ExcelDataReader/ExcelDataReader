using System;
using NUnit.Framework;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestCleanup = NUnit.Framework.TearDownAttribute;
using TestInitialize = NUnit.Framework.SetUpAttribute;
using TestMethod = NUnit.Framework.TestAttribute;

namespace ExcelDataReader.Tests
{
    [TestClass]
    public class ExcelReaderFactoryTests
    {
        [TestMethod]
        public void ProbeXls()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xls")))
            {
                Assert.AreEqual(excelReader.GetType().Name, "ExcelBinaryReader");
            }

            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestUnicodeChars.xls")))
            {
                Assert.AreEqual(excelReader.GetType().Name, "ExcelBinaryReader");
            }

            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("biff3.xls")))
            {
                Assert.AreEqual(excelReader.GetType().Name, "ExcelBinaryReader");
            }

            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("as3xls_BIFF2.xls")))
            {
                Assert.AreEqual(excelReader.GetType().Name, "ExcelBinaryReader");
            }
        }

        [TestMethod]
        public void ProbeXlsx()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xlsx")))
            {
                Assert.AreEqual(excelReader.GetType().Name, "ExcelOpenXmlReader");
            }

            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestOpenXml.xlsx")))
            {
                Assert.AreEqual(excelReader.GetType().Name, "ExcelOpenXmlReader");
            }
        }
    }
}
