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
        public void ProbeXLS()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10")))
            {
                Assert.AreEqual(excelReader.GetType(), typeof(ExcelBinaryReader));
            }

            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestUnicodeChars")))
            {
                Assert.AreEqual(excelReader.GetType(), typeof(ExcelBinaryReader));
            }

            using (var stream = Configuration.GetTestWorkbook("biff3"))
                Assert.Throws<NotSupportedException>(() => ExcelReaderFactory.CreateReader(stream));

        }

        [TestMethod]
        public void ProbeXLSX()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("xTest10x10")))
            {
                Assert.AreEqual(excelReader.GetType(), typeof(ExcelOpenXmlReader));
            }

            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("xTestOpenXml")))
            {
                Assert.AreEqual(excelReader.GetType(), typeof(ExcelOpenXmlReader));
            }
        }
    }
}
