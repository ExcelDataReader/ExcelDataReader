using System;
using NUnit.Framework;

namespace ExcelDataReader.Tests
{
    [TestFixture]
    public class ExcelReaderFactoryTests
    {
        [TestCase("Test10x10.xls")]
        [TestCase("TestUnicodeChars.xls")]
        [TestCase("biff3.xls")]
        [TestCase("as3xls_BIFF2.xls")]
        public void ProbeXls(string name)
        {
            using IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook(name));
            Assert.AreEqual(excelReader.GetType().Name, "ExcelBinaryReader");
        }

        [TestCase("Test10x10.xlsx")]
        [TestCase("TestOpen.xlsx")]
        [TestCase("TestOpen.xlsb")]
        public void ProbeOpenXml(string name)
        {
            using IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook(name));
            Assert.AreEqual(excelReader.GetType().Name, "ExcelOpenXmlReader");
        }
    }
}
