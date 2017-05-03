using System.Globalization;
using System.Threading;
using System;
using System.IO;
using NUnit.Framework;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestInitialize = NUnit.Framework.SetUpAttribute;
using TestCleanup = NUnit.Framework.TearDownAttribute;
using TestMethod = NUnit.Framework.TestAttribute;
using Excel;

namespace ExcelDataReader.Tests
{
    [TestClass]
    public class ExcelOpenXmlReaderLocaleTest
    {
        [TestMethod]
        public void Time_is_readable_for_polish_locale_issue_xxx()
        {
#if NET20 || NET45
			Thread.CurrentThread.CurrentCulture = new CultureInfo("pl-PL", false);

            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_xxx_LocaleTime")))
            {
                var dataset = reader.AsDataSet();

                Assert.AreEqual(new System.DateTime(1899, 12, 31, 1, 34, 0), dataset.Tables[0].Rows[1][1]);
                Assert.AreEqual(new System.DateTime(1899, 12, 31, 1, 34, 0), dataset.Tables[0].Rows[2][1]);
                Assert.AreEqual(new System.DateTime(1899, 12, 31, 18, 47, 0), dataset.Tables[0].Rows[3][1]);

                reader.Close();
            }
#endif
		}

		[TestMethod]
        public void Test_Decimal_Locale()
        {
#if NET20 || NET45
			//change culture to german. this will expect commas instead of decimal points
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE", false);

            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Decimal_Locale"));

            var dataSet = excelReader.AsDataSet();

            excelReader.Close();

            Assert.AreEqual(0.01, dataSet.Tables[0].Rows[0][0]);
            Assert.AreEqual(0.0001, dataSet.Tables[0].Rows[1][0]);
            Assert.AreEqual(0.123456789, dataSet.Tables[0].Rows[2][0]);
            Assert.AreEqual(0.00000000001, dataSet.Tables[0].Rows[3][0]);
#endif
		}
	}
}
