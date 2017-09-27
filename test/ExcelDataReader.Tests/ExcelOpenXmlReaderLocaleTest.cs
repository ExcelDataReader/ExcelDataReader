#if NET20 || NET45 || NETCOREAPP2_0
using System.Data;
using System.Globalization;
using System.Threading;
using NUnit.Framework;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestMethod = NUnit.Framework.TestAttribute;

// ReSharper disable InconsistentNaming
namespace ExcelDataReader.Tests
{
    [TestClass]
    public class ExcelOpenXmlReaderLocaleTest
    {
        [TestMethod]
        public void Time_is_readable_for_polish_locale_issue_xxx()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("pl-PL", false);

            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_xxx_LocaleTime")))
            {
                var dataset = reader.AsDataSet();

                Assert.AreEqual(new System.DateTime(1899, 12, 31, 1, 34, 0), dataset.Tables[0].Rows[1][1]);
                Assert.AreEqual(new System.DateTime(1899, 12, 31, 1, 34, 0), dataset.Tables[0].Rows[2][1]);
                Assert.AreEqual(new System.DateTime(1899, 12, 31, 18, 47, 0), dataset.Tables[0].Rows[3][1]);

                reader.Close();
            }
        }

        [TestMethod]
        public void Test_Decimal_Locale()
        {
            // change culture to german. this will expect commas instead of decimal points
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE", false);

            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Decimal_Locale"));

            var dataSet = excelReader.AsDataSet();

            excelReader.Close();

            Assert.AreEqual(0.01, dataSet.Tables[0].Rows[0][0]);
            Assert.AreEqual(0.0001, dataSet.Tables[0].Rows[1][0]);
            Assert.AreEqual(0.123456789, dataSet.Tables[0].Rows[2][0]);
            Assert.AreEqual(0.00000000001, dataSet.Tables[0].Rows[3][0]);
        }

        [TestMethod]
        // [SetCulture("sv-SE")]
        public void CellFormat49()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("sv-SE", false);

            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Format49_@")))
            {
                DataSet result = excelReader.AsDataSet();

                // ExcelDataReader used to convert numbers formatted with NumFmtId=49/@ to culture-specific strings.
                // This behaviour changed in v3 to return the original value:
                // Assert.That(result.Tables[0].Rows[0].ItemArray, Is.EqualTo(new[] { "2010-05-05", "1.1", "2,2", "123", "2,2" }));
                Assert.That(result.Tables[0].Rows[0].ItemArray, Is.EqualTo(new object[] { "2010-05-05", "1.1", 2.2000000000000002D, 123.0D, "2,2" }));
            }
        }
    }
}
#endif
