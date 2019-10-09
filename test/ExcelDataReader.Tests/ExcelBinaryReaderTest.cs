using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader.Exceptions;

using NUnit.Framework;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestCleanup = NUnit.Framework.TearDownAttribute;
using TestInitialize = NUnit.Framework.SetUpAttribute;
using TestMethod = NUnit.Framework.TestAttribute;

namespace ExcelDataReader.Tests
{
    [TestClass]

    public class ExcelBinaryReaderTest
    {
        [OneTimeSetUp]
        public void TestInitialize()
        {
#if NETCOREAPP1_0 || NETCOREAPP2_0
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif
        }

        [TestMethod]
        public void GitIssue70ExcelBinaryReaderTryConvertOADateTimeFormula()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_70_ExcelBinaryReader_tryConvertOADateTime _convert_dates.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);

                var date = ds.Tables[0].Rows[1].ItemArray[0];

                Assert.AreEqual(new DateTime(2014, 01, 01), date);
            }
        }

        [TestMethod]
        public void GitIssue51ReadCellLabel()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_51.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);

                var value = ds.Tables[0].Rows[0].ItemArray[1];

                Assert.AreEqual("Monetary aggregates (R millions)", value);
            }
        }

        [TestMethod]
        public void GitIssue29ReadSheetStatesReadsCorrectly()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Excel_Dataset.xls")))
            {
                Assert.AreEqual("hidden", excelReader.VisibleState);

                excelReader.NextResult();
                Assert.AreEqual("visible", excelReader.VisibleState);

                excelReader.NextResult();
                Assert.AreEqual("veryhidden", excelReader.VisibleState);
            }
        }

        [TestMethod]
        public void GitIssue29AsDataSetProvidesCorrectSheetVisibleState()
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Excel_Dataset.xls")))
            {
                var dataSet = reader.AsDataSet();

                Assert.IsTrue(dataSet != null);
                Assert.AreEqual(3, dataSet.Tables.Count);
                Assert.AreEqual("hidden", dataSet.Tables[0].ExtendedProperties["visiblestate"]);
                Assert.AreEqual("visible", dataSet.Tables[1].ExtendedProperties["visiblestate"]);
                Assert.AreEqual("veryhidden", dataSet.Tables[2].ExtendedProperties["visiblestate"]);
            }
        }

        [TestMethod]
        public void GitIssue389FilterSheetByVisibility()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Excel_Dataset.xls")))
            {
                var result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    FilterSheet = (r, index) => r.VisibleState == "visible"
                });

                Assert.AreEqual(1, result.Tables.Count);
            }
        }

        [TestMethod]
        public void GitIssue45()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_45.xls")))
            {
                do
                {
                    while (reader.Read())
                    {
                    }
                }
                while (reader.NextResult());
            }
        }

        [TestMethod]
        public void AsDataSetTest()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestChess.xls")))
            {
                DataSet result = excelReader.AsDataSet();

                Assert.IsTrue(result != null);
                Assert.AreEqual(1, result.Tables.Count);
                Assert.AreEqual(4, result.Tables[0].Rows.Count);
                Assert.AreEqual(6, result.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void AsDataSetTestRowCount()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestChess.xls")))
            {
                var result = excelReader.AsDataSet(Configuration.NoColumnNamesConfiguration);

                Assert.AreEqual(4, result.Tables[0].Rows.Count);
            }
        }

        [TestMethod]
        public void AsDataSetTestRowCountFirstRowAsColumnNames()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestChess.xls")))
            {
                var result = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(3, result.Tables[0].Rows.Count);
            }
        }

        [TestMethod]
        public void Issue1155311570FatIssueOffset()
        {
            void DoTestFatStreamIssue(string sheetId)
            {
                string filePath;
                using (var excelReader1 = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook(sheetId))) // Works.
                {
                    filePath = Configuration.GetTestWorkbookPath(sheetId);
                    Assert.IsNotNull(excelReader1);
                }

                using (var ms1 = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var excelReader2 = ExcelReaderFactory.CreateBinaryReader(ms1)) // Works!
                    Assert.IsNotNull(excelReader2);

                var bytes = File.ReadAllBytes(filePath);
                using (var ms2 = new MemoryStream(bytes))
                using (var excelReader3 = ExcelReaderFactory.CreateBinaryReader(ms2)) // Did not work, but does now
                    Assert.IsNotNull(excelReader3);
            }

            void DoTestFatStreamIssueType2(string sheetId)
            {
                var filePath = Configuration.GetTestWorkbookPath(sheetId);

                using (Stream stream = new MemoryStream(File.ReadAllBytes(filePath)))
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream))
                {
                    // ReSharper disable once AccessToDisposedClosure
                    Assert.DoesNotThrow(() => excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration));
                }
            }

            DoTestFatStreamIssue("Test_Issue_11553_FAT.xls");
            DoTestFatStreamIssueType2("Test_Issue_11570_FAT_1.xls");
            DoTestFatStreamIssueType2("Test_Issue_11570_FAT_2.xls");
        }

        /*[TestMethod]
        public void Test_SSRS()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_SSRS"));
            DataSet result = excelReader.AsDataSet();
            excelReader.Close();
        }*/

        [TestMethod]
        public void ChessTest()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestChess.xls")))
            {
                DataTable result = excelReader.AsDataSet().Tables[0];

                Assert.AreEqual(4, result.Rows.Count);
                Assert.AreEqual(6, result.Columns.Count);
                Assert.AreEqual("1", result.Rows[3][5].ToString());
                Assert.AreEqual("1", result.Rows[2][0].ToString());
            }
        }

        [TestMethod]
        public void DataReaderNextResultTest()
        {
            using (IExcelDataReader r = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestMultiSheet.xls")))
            {
                Assert.AreEqual(3, r.ResultsCount);

                var table = new DataTable();
                table.Columns.Add("c1", typeof(int));
                table.Columns.Add("c2", typeof(int));
                table.Columns.Add("c3", typeof(int));
                table.Columns.Add("c4", typeof(int));

                int fieldCount = -1;

                while (r.Read())
                {
                    fieldCount = r.FieldCount;
                    table.Rows.Add(
                        Convert.ToInt32(r.GetValue(0)),
                        Convert.ToInt32(r.GetValue(1)),
                        Convert.ToInt32(r.GetValue(2)),
                        Convert.ToInt32(r.GetValue(3)));
                }

                Assert.AreEqual(12, table.Rows.Count);
                Assert.AreEqual(12, r.RowCount);
                Assert.AreEqual(4, fieldCount);
                Assert.AreEqual(1, table.Rows[11][3]);

                r.NextResult();
                table.Rows.Clear();

                while (r.Read())
                {
                    fieldCount = r.FieldCount;
                    table.Rows.Add(
                        Convert.ToInt32(r.GetValue(0)),
                        Convert.ToInt32(r.GetValue(1)),
                        Convert.ToInt32(r.GetValue(2)),
                        Convert.ToInt32(r.GetValue(3)));
                }

                Assert.AreEqual(12, table.Rows.Count);
                Assert.AreEqual(12, r.RowCount);
                Assert.AreEqual(4, fieldCount);
                Assert.AreEqual(2, table.Rows[11][3]);

                r.NextResult();
                table.Rows.Clear();

                while (r.Read())
                {
                    fieldCount = r.FieldCount;
                    table.Rows.Add(
                        Convert.ToInt32(r.GetValue(0)),
                        Convert.ToInt32(r.GetValue(1)));
                }

                Assert.AreEqual(5, table.Rows.Count);
                Assert.AreEqual(5, r.RowCount);
                Assert.AreEqual(2, fieldCount);
                Assert.AreEqual(3, table.Rows[4][1]);

                Assert.AreEqual(false, r.NextResult());
            }
        }

        [TestMethod]
        public void DataReaderReadTest()
        {
            using (IExcelDataReader r = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_num_double_date_bool_string.xls")))
            {
                var table = new DataTable();
                table.Columns.Add(new DataColumn("num_col", typeof(int)));
                table.Columns.Add(new DataColumn("double_col", typeof(double)));
                table.Columns.Add(new DataColumn("date_col", typeof(DateTime)));
                table.Columns.Add(new DataColumn("boo_col", typeof(bool)));

                int fieldCount = -1;

                while (r.Read())
                {
                    fieldCount = r.FieldCount;
                    table.Rows.Add(
                        Convert.ToInt32(r.GetValue(0)),
                        Convert.ToDouble(r.GetValue(1)),
                        r.GetDateTime(2),
                        r.IsDBNull(4));
                }

                Assert.AreEqual(6, fieldCount);

                Assert.AreEqual(30, table.Rows.Count);

                Assert.AreEqual(1, int.Parse(table.Rows[0][0].ToString()));
                Assert.AreEqual(1346269, int.Parse(table.Rows[29][0].ToString()));

                // double + Formula
                Assert.AreEqual(1.02, double.Parse(table.Rows[0][1].ToString()));
                Assert.AreEqual(4.08, double.Parse(table.Rows[2][1].ToString()));
                Assert.AreEqual(547608330.24, double.Parse(table.Rows[29][1].ToString()));

                // Date + Formula
                Assert.AreEqual(new DateTime(2009, 5, 11).ToShortDateString(), DateTime.Parse(table.Rows[0][2].ToString()).ToShortDateString());
                Assert.AreEqual(new DateTime(2009, 11, 30).ToShortDateString(), DateTime.Parse(table.Rows[29][2].ToString()).ToShortDateString());
            }
        }

        [TestMethod]
        public void Dimension10X10000Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test10x10000.xls")))
            {
                DataTable result = excelReader.AsDataSet().Tables[0];

                Assert.AreEqual(10000, result.Rows.Count);
                Assert.AreEqual(10, result.Columns.Count);
                Assert.AreEqual("1x2", result.Rows[1][1]);
                Assert.AreEqual("1x10", result.Rows[1][9]);
                Assert.AreEqual("1x1", result.Rows[9999][0]);
                Assert.AreEqual("1x10", result.Rows[9999][9]);
            }
        }

        [TestMethod]
        public void Dimension10X10Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test10x10.xls")))
            {
                DataTable result = excelReader.AsDataSet().Tables[0];

                Assert.AreEqual(10, result.Rows.Count);
                Assert.AreEqual(10, result.Columns.Count);
                Assert.AreEqual("10x10", result.Rows[1][0]);
                Assert.AreEqual("10x27", result.Rows[9][9]);
            }
        }

        [TestMethod]
        public void Dimension255X10Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test255x10.xls")))
            {
                DataTable result = excelReader.AsDataSet().Tables[0];

                Assert.AreEqual(10, result.Rows.Count);
                Assert.AreEqual(255, result.Columns.Count);
                Assert.AreEqual("1", result.Rows[9][254].ToString());
                Assert.AreEqual("one", result.Rows[1][1].ToString());
            }
        }

        [TestMethod]
        public void DoublePrecisionTest()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestDoublePrecision.xls")))
            {
                DataTable result = excelReader.AsDataSet().Tables[0];

                const double excelPi = 3.14159265358979;

                Assert.AreEqual(+excelPi, result.Rows[2][1]);
                Assert.AreEqual(-excelPi, result.Rows[3][1]);

                Assert.AreEqual(+excelPi * 1.0e-300, result.Rows[4][1]);
                Assert.AreEqual(-excelPi * 1.0e-300, result.Rows[5][1]);

                Assert.AreEqual(+excelPi * 1.0e300, (double)result.Rows[6][1], 1e286); // only accurate to 1e286 because excel only has 15 digits precision
                Assert.AreEqual(-excelPi * 1.0e300, (double)result.Rows[7][1], 1e286);

                Assert.AreEqual(+excelPi * 1.0e14, result.Rows[8][1]);
                Assert.AreEqual(-excelPi * 1.0e14, result.Rows[9][1]);
            }
        }

        [TestMethod]
        public void FailTest()
        {
            var exception = Assert.Throws<HeaderException>(() =>
            {
                using (ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestFail_Binary.xls")))
                {
                }
            });

            Assert.AreEqual("Invalid file signature.", exception.Message);
        }

        [TestMethod]
        public void IssueDateAndTime1468Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Encoding_Formula_Date_1520.xls")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                string val1 = new DateTime(2009, 05, 01).ToShortDateString();
                string val2 = DateTime.Parse(dataSet.Tables[0].Rows[1][1].ToString()).ToShortDateString();

                Assert.AreEqual(val1, val2);

                val1 = DateTime.Parse("11:00:00").ToShortTimeString();
                val2 = DateTime.Parse(dataSet.Tables[0].Rows[2][4].ToString()).ToShortTimeString();

                Assert.AreEqual(val1, val2);
            }
        }

        [TestMethod]
        public void Issue8536Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_8536.xls")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                // date
                var date1900 = dataSet.Tables[0].Rows[7][1];
                Assert.AreEqual(typeof(DateTime), date1900.GetType());
                Assert.AreEqual(new DateTime(1900, 1, 1), date1900);

                // xml encoded chars
                var xmlChar1 = dataSet.Tables[0].Rows[6][1];
                Assert.AreEqual(typeof(string), xmlChar1.GetType());
                Assert.AreEqual("&#x26; ", xmlChar1);

                // number but matches a date serial
                var dateNum1 = dataSet.Tables[0].Rows[5][1];
                Assert.AreEqual(typeof(double), dateNum1.GetType());
                Assert.AreEqual(41244, double.Parse(dateNum1.ToString()));

                // date
                var date1 = dataSet.Tables[0].Rows[4][1];
                Assert.AreEqual(typeof(DateTime), date1.GetType());
                Assert.AreEqual(new DateTime(2012, 12, 1), date1);

                // double
                var num1 = dataSet.Tables[0].Rows[3][1];
                Assert.AreEqual(typeof(double), num1.GetType());
                Assert.AreEqual(12345, double.Parse(num1.ToString()));

                // boolean issue
                var val2 = dataSet.Tables[0].Rows[2][1];
                Assert.AreEqual(typeof(bool), val2.GetType());
                Assert.IsTrue((bool)val2);
            }
        }

        [TestMethod]
        public void Issue11397CurrencyTest()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11397.xls")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual("$44.99", dataSet.Tables[0].Rows[1][0].ToString()); // general in spreadsheet so should be a string including the $
                Assert.AreEqual(44.99, double.Parse(dataSet.Tables[0].Rows[2][0].ToString())); // currency euros in spreadsheet so should be a currency
                Assert.AreEqual(44.99, double.Parse(dataSet.Tables[0].Rows[3][0].ToString())); // currency pounds in spreadsheet so should be a currency
            }
        }

        [TestMethod]
        public void Issue4031NullColumn()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_4031_NullColumn.xls")))
            {
                // DataSet dataSet = excelReader.AsDataSet(true);
                excelReader.Read();
                Assert.IsNull(excelReader.GetValue(0));
                Assert.AreEqual("a", excelReader.GetString(1));
                Assert.AreEqual("b", excelReader.GetString(2));
                Assert.IsNull(excelReader.GetValue(3));
                Assert.AreEqual("d", excelReader.GetString(4));

                excelReader.Read();
                Assert.IsNull(excelReader.GetValue(0));
                Assert.IsNull(excelReader.GetValue(1));
                Assert.AreEqual("Test", excelReader.GetString(2));
                Assert.IsNull(excelReader.GetValue(3));
                Assert.AreEqual(1, excelReader.GetDouble(4));
            }
        }

        [TestMethod]
        public void Issue10725()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_10725.xls")))
            {
                excelReader.Read();
                Assert.AreEqual(8.8, excelReader.GetValue(0));

                DataSet result = excelReader.AsDataSet();

                Assert.AreEqual(8.8, result.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void Issue11435Colors()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11435_Colors.xls")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual("test1", dataSet.Tables[0].Rows[0][0].ToString());
                Assert.AreEqual("test2", dataSet.Tables[0].Rows[0][1].ToString());
                Assert.AreEqual("test3", dataSet.Tables[0].Rows[0][2].ToString());

                excelReader.Read();

                Assert.AreEqual("test1", excelReader.GetString(0));
                Assert.AreEqual("test2", excelReader.GetString(1));
                Assert.AreEqual("test3", excelReader.GetString(2));
            }
        }

        [TestMethod]
        public void Issue7433IllegalOleAutDate()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_7433_IllegalOleAutDate.xls")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual(3.10101195608231E+17, dataSet.Tables[0].Rows[0][0]);
                Assert.AreEqual("B221055625", dataSet.Tables[0].Rows[1][0]);
                Assert.AreEqual(4.12721197309241E+17, dataSet.Tables[0].Rows[2][0]);
            }
        }

        [TestMethod]
        public void IssueBoolFormula()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_BoolFormula.xls")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual(true, dataSet.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void IssueDecimal1109Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Decimal_1109.xls")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual(3.14159, dataSet.Tables[0].Rows[0][0]);

                const double val1 = -7080.61;
                double val2 = (double)dataSet.Tables[0].Rows[0][1];
                Assert.AreEqual(val1, val2);
            }
        }

        [TestMethod]
        public void IssueEncoding1520Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Encoding_Formula_Date_1520.xls")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                string val1 = "Simon Hodgetts";
                string val2 = dataSet.Tables[0].Rows[2][0].ToString();
                Assert.AreEqual(val1, val2);

                val1 = "John test";
                val2 = dataSet.Tables[0].Rows[1][0].ToString();
                Assert.AreEqual(val1, val2);

                // librement réutilisable
                val1 = "librement réutilisable";
                val2 = dataSet.Tables[0].Rows[7][0].ToString();
                Assert.AreEqual(val1, val2);

                val2 = dataSet.Tables[0].Rows[8][0].ToString();
                Assert.AreEqual(val1, val2);
            }
        }

        [TestMethod]
        public void MultiSheetTest()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestMultiSheet.xls")))
            {
                DataSet result = excelReader.AsDataSet();

                Assert.AreEqual(3, result.Tables.Count);

                Assert.AreEqual(4, result.Tables["Sheet1"].Columns.Count);
                Assert.AreEqual(12, result.Tables["Sheet1"].Rows.Count);
                Assert.AreEqual(4, result.Tables["Sheet2"].Columns.Count);
                Assert.AreEqual(12, result.Tables["Sheet2"].Rows.Count);
                Assert.AreEqual(2, result.Tables["Sheet3"].Columns.Count);
                Assert.AreEqual(5, result.Tables["Sheet3"].Rows.Count);

                Assert.AreEqual("1", result.Tables["Sheet2"].Rows[11][0].ToString());
                Assert.AreEqual("2", result.Tables["Sheet1"].Rows[11][3].ToString());
                Assert.AreEqual("3", result.Tables["Sheet3"].Rows[4][1].ToString());
            }
        }

        [TestMethod]
        public void TestNumDoubleDateBoolString()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_num_double_date_bool_string.xls")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual(30, dataSet.Tables[0].Rows.Count);
                Assert.AreEqual(6, dataSet.Tables[0].Columns.Count);

                Assert.AreEqual(1, int.Parse(dataSet.Tables[0].Rows[0][0].ToString()));
                Assert.AreEqual(1346269, int.Parse(dataSet.Tables[0].Rows[29][0].ToString()));

                // bool
                Assert.IsNotNull(dataSet.Tables[0].Rows[22][3].ToString());
                Assert.AreEqual(dataSet.Tables[0].Rows[22][3], true);

                // double + Formula
                Assert.AreEqual(1.02, double.Parse(dataSet.Tables[0].Rows[0][1].ToString()));
                Assert.AreEqual(4.08, double.Parse(dataSet.Tables[0].Rows[2][1].ToString()));
                Assert.AreEqual(547608330.24, double.Parse(dataSet.Tables[0].Rows[29][1].ToString()));

                // Date + Formula
                Assert.AreEqual(new DateTime(2009, 5, 11), dataSet.Tables[0].Rows[0][2]);
                Assert.AreEqual(new DateTime(2009, 11, 30), dataSet.Tables[0].Rows[29][2]);

                // Custom Date Time + Formula
                var s = dataSet.Tables[0].Rows[0][5].ToString();
                Assert.AreEqual(new DateTime(2009, 5, 7, 11, 1, 2), DateTime.Parse(s));
                s = dataSet.Tables[0].Rows[1][5].ToString();
                Assert.AreEqual(new DateTime(2009, 5, 8, 11, 1, 2), DateTime.Parse(s));

                // DBNull value
                Assert.AreEqual(DBNull.Value, dataSet.Tables[0].Rows[1][4]);
            }
        }

        [TestMethod]
        public void Issue11479BlankSheet()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11479_BlankSheet.xls")))
            {
                // DataSet result = excelReader.AsDataSet(true);
                excelReader.Read();
                Assert.AreEqual(5, excelReader.FieldCount);
                excelReader.NextResult();
                excelReader.Read();
                Assert.AreEqual(0, excelReader.FieldCount);

                excelReader.NextResult();
                excelReader.Read();
                Assert.AreEqual(0, excelReader.FieldCount);
            }
        }

        [TestMethod]
        public void TestBlankHeader()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_BlankHeader.xls")))
            {
                excelReader.Read();
                Assert.AreEqual(6, excelReader.FieldCount);
                excelReader.Read();
                for (int i = 0; i < excelReader.FieldCount; i++)
                {
                    Console.WriteLine("{0}:{1}", i, excelReader.GetValue(i));
                }
            }
        }

        [TestMethod]
        public void TestOpenOffice()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_OpenOffice.xls")))
            {
                AssertUtilities.DoOpenOfficeTest(excelReader);
            }
        }

        /// <summary>
        /// Issue 11 - OpenOffice files were skipping the first row if IsFirstRowAsColumnNames = false;
        /// </summary>
        [TestMethod]
        public void GitIssue11OpenOfficeRowCount()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_OpenOffice.xls")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.NoColumnNamesConfiguration);
                Assert.AreEqual(34, dataSet.Tables[0].Rows.Count);
            }
        }

        /// <summary>
        /// This test is to ensure that we get the same results from an xls saved in excel vs open office
        /// </summary>
        [TestMethod]
        public void TestOpenOfficeSavedInExcel()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Excel_OpenOffice.xls")))
            {
                AssertUtilities.DoOpenOfficeTest(excelReader);
            }
        }

        [TestMethod]
        public void TestIssue11601ReadSheetNames()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Excel_Dataset.xls")))
            {
                Assert.AreEqual("test.csv", excelReader.Name);

                excelReader.NextResult();
                Assert.AreEqual("Sheet2", excelReader.Name);

                excelReader.NextResult();
                Assert.AreEqual("Sheet3", excelReader.Name);
            }
        }

        [TestMethod]
        public void UnicodeCharsTest()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestUnicodeChars.xls")))
            {
                DataTable result = excelReader.AsDataSet().Tables[0];

                Assert.AreEqual(3, result.Rows.Count);
                Assert.AreEqual(8, result.Columns.Count);
                Assert.AreEqual("\u00e9\u0417", result.Rows[1][0].ToString());
            }
        }

        [TestMethod]
        public void UncalculatedTest()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Uncalculated.xls")))
            {
                var dataSet = excelReader.AsDataSet();
                Assert.IsNotNull(dataSet);
                Assert.AreNotEqual(dataSet.Tables.Count, 0);
                var table = dataSet.Tables[0];
                Assert.IsNotNull(table);

                Assert.AreEqual("1", table.Rows[1][0].ToString());
                Assert.AreEqual("3", table.Rows[1][2].ToString());
                Assert.AreEqual("3", table.Rows[1][4].ToString());
            }
        }

        [TestMethod]
        public void Issue11570Excel2013()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11570_Excel2013.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual(2, dataSet.Tables[0].Columns.Count);
                Assert.AreEqual(5, dataSet.Tables[0].Rows.Count);

                Assert.AreEqual("1.1.1.2", dataSet.Tables[0].Rows[0][0]);
                Assert.AreEqual(10d, dataSet.Tables[0].Rows[0][1]);

                Assert.AreEqual("1.1.1.15", dataSet.Tables[0].Rows[1][0]);
                Assert.AreEqual(3d, dataSet.Tables[0].Rows[1][1]);

                Assert.AreEqual("2.1.2.23", dataSet.Tables[0].Rows[2][0]);
                Assert.AreEqual(14d, dataSet.Tables[0].Rows[2][1]);

                Assert.AreEqual("2.1.2.31", dataSet.Tables[0].Rows[3][0]);
                Assert.AreEqual(2d, dataSet.Tables[0].Rows[3][1]);

                Assert.AreEqual("2.8.7.30", dataSet.Tables[0].Rows[4][0]);
                Assert.AreEqual(2d, dataSet.Tables[0].Rows[4][1]);
            }
        }

        [TestMethod]
        public void Issue11572CodePage()
        {
            // This test was skipped for a long time as it produced: "System.NotSupportedException : No data is available for encoding 27651."
            // Upon revisiting the underlying cause appears to be fixed
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11572_CodePage.xls")))
            {
                Assert.DoesNotThrow(() => excelReader.AsDataSet());
            }
        }

        /// <summary>
        /// Not fixed yet
        /// </summary>
        [TestMethod]
        public void Issue11545NoIndex()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11545_NoIndex.xls")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual("CI2229         ", dataSet.Tables[0].Rows[0][0]);
                Assert.AreEqual("12069E01018A1  ", dataSet.Tables[0].Rows[0][6]);
                Assert.AreEqual(new DateTime(2012, 03, 01), dataSet.Tables[0].Rows[0][8]);
            }
        }

        [TestMethod]
        public void Issue11573BlankValues()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11573_BlankValues.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual(1D, dataSet.Tables[0].Rows[12][0]);
                Assert.AreEqual("070202", dataSet.Tables[0].Rows[12][1]);
            }
        }

        [TestMethod]
        public void IssueDateFormatButNotDate()
        {
            // we want to make sure that if a cell is formatted as a date but it's contents are not a date then
            // the output is not a date
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_DateFormatButNotDate.xls")))
            {
                excelReader.Read();
                Assert.AreEqual("columna", excelReader.GetValue(0));
                Assert.AreEqual("columnb", excelReader.GetValue(1));
                Assert.AreEqual("columnc", excelReader.GetValue(2));
                Assert.AreEqual("columnd", excelReader.GetValue(3));
                Assert.AreEqual("columne", excelReader.GetValue(4));

                excelReader.Read();
                Assert.AreEqual(1D, excelReader.GetValue(0));
                Assert.AreEqual(2D, excelReader.GetValue(1));
                Assert.AreEqual(3D, excelReader.GetValue(2));
                var value = excelReader.GetValue(3);
                Assert.AreEqual(new DateTime(2013, 12, 10), value);
                Assert.AreEqual("red", excelReader.GetValue(4));

                excelReader.Read();
                Assert.AreEqual("yellow", excelReader.GetValue(4));
            }
        }

        [TestMethod]
        public void Issue11642ValuesNotLoaded()
        {
            // Excel.Log.Log.InitializeWith<Log4NetLog>();
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11642_ValuesNotLoaded.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual("431113*", dataSet.Tables[2].Rows[29][1].ToString());
                Assert.AreEqual("024807", dataSet.Tables[2].Rows[36][1].ToString());
                Assert.AreEqual("160019", dataSet.Tables[2].Rows[53][1].ToString());
            }
        }

        [TestMethod]
        public void Issue11636BiffStream()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11636_BiffStream.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                // check a couple of values
                Assert.AreEqual("SP011", dataSet.Tables[0].Rows[9][0]);
                Assert.AreEqual(9.9, dataSet.Tables[0].Rows[32][11]);
                Assert.AreEqual(78624.44, dataSet.Tables[1].Rows[27][12]);
            }
        }

        /// <summary>
        /// Not fixed yet
        /// The problem occurs with unseekable stream and logic related to minifat that uses seek
        /// It should probably only use seek if it needs to go backwards, I think at the moment it uses seek all the time
        /// which is probably not good for performance
        /// </summary>
        [TestMethod]
        [Ignore("Not fixed yet")]
        public void Issue1163911644ForwardOnlyStream()
        {
            // Excel.Log.Log.InitializeWith<Log4NetLog>();
            using (var stream = Configuration.GetTestWorkbook("Test_OpenOffice"))
            {
                using (var forwardStream = SeekErrorMemoryStream.CreateFromStream(stream))
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(forwardStream))
                {
                    Assert.DoesNotThrow(() => excelReader.AsDataSet());
                }
            }
        }

        /// <summary>
        /// Not fixed yet
        /// The problem occurs with unseekable stream and logic related to minifat that uses seek
        /// It should probably only use seek if it needs to go backwards, I think at the moment it uses seek all the time
        /// which is probably not good for performance
        /// </summary>
        [TestMethod]
        public void Issue12556Corrupt()
        {
            Assert.Throws<CompoundDocumentException>(() =>
            {
                // Excel.Log.Log.InitializeWith<Log4NetLog>();
                using (var forwardStream = Configuration.GetTestWorkbook("Test_Issue_12556_corrupt.xls"))
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(forwardStream))
                {
                    Assert.DoesNotThrow(() => excelReader.AsDataSet());
                }
            });
        }

        /// <summary>
        /// Some spreadsheets were crashing with index out of range error (from SSRS)
        /// </summary>
        [TestMethod]
        public void TestIssue11818OutOfRange()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11818_OutOfRange.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual("Total Revenue", dataSet.Tables[0].Rows[10][0]);
            }
        }

        [TestMethod]
        public void TestIssue111NoRowRecords()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_111_NoRowRecords.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual(1, dataSet.Tables.Count);
                Assert.AreEqual(12, dataSet.Tables[0].Rows.Count);
                Assert.AreEqual(14, dataSet.Tables[0].Columns.Count);

                Assert.AreEqual(2015.0, dataSet.Tables[0].Rows[7][0]);
            }
        }

        [TestMethod]
        public void TestGitIssue145()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_145.xls")))
            {
                excelReader.Read();
                excelReader.Read();
                excelReader.Read();

                string value = excelReader.GetString(3);

                Assert.AreEqual("Japanese Government Bonds held by the Bank of Japan", value);
            }
        }

        [TestMethod]
        public void TestGitIssue152SheetNameUtf16LeCompressed()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_152.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual("åäöñ", dataSet.Tables[0].TableName);
            }
        }

        [TestMethod]
        public void TestGitIssue152CellUtf16LeCompressed()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_152.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual("åäöñ", dataSet.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void GitIssue158()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_158.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);

                var date = ds.Tables[0].Rows[3].ItemArray[2];

                Assert.AreEqual(new DateTime(2016, 09, 10), date);
            }
        }

        [TestMethod]
        public void GitIssue173()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_173.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);
                Assert.AreEqual(40, ds.Tables.Count);
            }
        }

        [TestMethod]
        public void ReadWriteProtectedStructureUsingStandardEncryption()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("protectedsheet-xxx.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);
                Assert.AreEqual("x", ds.Tables[0].Rows[0][0]);
                Assert.AreEqual(1.4, ds.Tables[0].Rows[1][0]);
            }
        }

        [TestMethod]
        public void TestIncludeTableWithOnlyImage()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestTableOnlyImage_x01oct2016.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);
                Assert.AreEqual(4, ds.Tables.Count);
            }
        }

        [TestMethod]
        public void AllowFfffAsByteOrder()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_InvalidByteOrderValueInHeader.xls")))
            {
                int tableCount = 0;
                do
                {
                    while (excelReader.Read())
                    {
                    }

                    tableCount++;
                }
                while (excelReader.NextResult());

                Assert.AreEqual(454, tableCount);
            }
        }

        [TestMethod]
        public void HandleRowBlocksWithOutOfOrderCells()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("AllColumnsNotReadInHiddenTable.xls")))
            {
                var ds = excelReader.AsDataSet();

                object[] expected = { "21/09/2015", 1187.5282349881188, 650.8582749049624, 1361.7209439645526, 321.74647548613916, 369.48879457369037 };

                Assert.AreEqual(51, ds.Tables[1].Rows.Count);
                Assert.AreEqual(expected, ds.Tables[1].Rows[1].ItemArray);
            }
        }

        [TestMethod]
        public void HandleRowBlocksWithDifferentNumberOfColumnsAndInvalidDimensions()
        {
            // http://www.ine.cl/canales/chile_estadistico/estadisticas_economicas/edificacion/archivos/xls/edificacion_totalpais_seriehistorica_enero_2017.xls
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("RowWithDifferentNumberOfColumns.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual(256, ds.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void IfNoDimensionDetermineFieldCountByProcessingAllCellColumnIndexes()
        {
            // This xls file has a row record with 256 columns but only values for 6.
            // This test was created when ExcelDataReader incorrectly dropped 8
            // bits off the dimensions' LastColumn in BIFF8 files and relied
            // on scanning to come up with 6 columns. The test was changed to
            // assume valid dimensions:
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_145.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual(256, ds.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void Row1217NotRead()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Row1217NotRead.xls")))
            {
                var ds = excelReader.AsDataSet();
                CollectionAssert.AreEqual(new object[] {
                    DBNull.Value,
                    "Año",
                    "Mes",
                    DBNull.Value, //Merged Cell
                    "Índice",
                    "Variación Mensual",
                    "Variación Acumulada",
                    "Variación en 12 Meses",
                    "Incidencia Mensual",
                    "Incidencia Acumulada", "" +
                    "Incidencia a 12 Meses",
                    DBNull.Value, //Merged Cell
                    DBNull.Value }, ds.Tables[0].Rows[1216].ItemArray);
            }
        }

        [TestMethod]
        public void StringContinuationAfterCharacterData()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("StringContinuationAfterCharacterData.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual("商業動態統計速報-平成29年2月分-  統計表", ds.Tables[0].Rows[3][2]);
                Assert.AreEqual("Preliminary Report on the Current Survey of Commerce  ( February,2017 )　Statistics Tables", ds.Tables[0].Rows[4][2]);
                Assert.AreEqual("\nWholesale", ds.Tables[1].Rows[18][9]);
            }
        }

        [TestCase]
        public void Biff3IsSupported()
        {
            using (var stream = Configuration.GetTestWorkbook("biff3.xls"))
            {
                using (var reader = ExcelReaderFactory.CreateBinaryReader(stream))
                {
                    reader.AsDataSet();
                }
            }
        }

        [TestCase]
        public void GitIssue5()
        {
            using (var stream = Configuration.GetTestWorkbook("Test_git_issue_5.xls"))
                Assert.Throws<CompoundDocumentException>(() => ExcelReaderFactory.CreateBinaryReader(stream));
        }

        [TestCase]
        public void Issue2InvalidDimensionRecord()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_2.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual(new[] { "A1", "B1" }, ds.Tables[0].Rows[0].ItemArray);
            }
        }

        [TestCase]
        public void ExcelLibraryNonContinuousMiniStream()
        {
            // Verify the output from the sample code for the ExcelLibrary package parses
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("ExcelLibrary_newdoc.xls")))
            {
                Assert.DoesNotThrow(() => excelReader.AsDataSet());
            }
        }

        [TestCase]
        public void GitIssue184AdditionalFatSectors()
        {
            // Big spreadsheets have additional sectors beyond the header with FAT contents
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("GitIssue_184_FATSectors.xls")))
            {
                DataSet ds = null;
                Assert.DoesNotThrow(() => ds = excelReader.AsDataSet());
                Assert.AreEqual(12, ds.Tables.Count);
                Assert.AreEqual("DATAS (12)", ds.Tables[0].TableName);
                Assert.AreEqual("DATAS (5)", ds.Tables[11].TableName);
            }
        }

        [TestMethod]
        public void RowContentSpreadOverMultipleBlocks()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_217.xls")))
            {
                var ds = excelReader.AsDataSet();
                CollectionAssert.AreEqual(new object[] { "REX GESAMT      ", 484.7929, 142.1032, -0.1656, 5.0315225293000001, 5.0398685515999997, 37.5344725251, DBNull.Value, DBNull.Value }, ds.Tables[2].Rows[10].ItemArray);
            }
        }

        [TestMethod]
        public void GitIssue231NoCodePage()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_231_NoCodePage.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual(11, ds.Tables[0].Columns.Count);
                Assert.AreEqual(5, ds.Tables[0].Rows.Count);
            }
        }

        [TestMethod]
        public void GitIssue82Date1900Binary()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("roo_1900_base.xls")))
            {
                // 15/06/2009
                // 28/06/2009 (=TODAY() when file was saved)

                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(new DateTime(2009, 6, 15), (DateTime)result.Tables[0].Rows[0][0]);
                Assert.AreEqual(new DateTime(2009, 6, 28), (DateTime)result.Tables[0].Rows[1][0]);
            }
        }

        [TestMethod]
        public void GitIssue82Date1904Binary()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("roo_1904_base.xls")))
            {
                // 15/06/2009
                // 28/06/2009 (=TODAY() when file was saved)

                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(new DateTime(2009, 6, 15), (DateTime)result.Tables[0].Rows[0][0]);
                Assert.AreEqual(new DateTime(2009, 6, 28), (DateTime)result.Tables[0].Rows[1][0]);
            }
        }

        [TestMethod]
        public void As3XlsBiff2()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF2.xls")))
            {
                DataSet result = excelReader.AsDataSet();
                TestAs3Xls(result);
            }
        }

        [TestMethod]
        public void As3XlsBiff3()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF3.xls")))
            {
                DataSet result = excelReader.AsDataSet();
                TestAs3Xls(result);
            }
        }

        [TestMethod]
        public void As3XlsBiff4()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF4.xls")))
            {
                DataSet result = excelReader.AsDataSet();
                TestAs3Xls(result);
            }
        }

        [TestMethod]
        public void As3XlsBiff5()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF5.xls")))
            {
                DataSet result = excelReader.AsDataSet();
                TestAs3Xls(result);
            }
        }

        private static void TestAs3Xls(DataSet result)
        {
            Assert.AreEqual(1, result.Tables[0].Rows[0][0]);
            Assert.AreEqual("Hi", result.Tables[0].Rows[0][1]);
            Assert.AreEqual(10.22D, result.Tables[0].Rows[0][2]);
            Assert.AreEqual(14.754317602356753D, result.Tables[0].Rows[0][3]);
            Assert.AreEqual(21.04107572533686D, result.Tables[0].Rows[0][4]);

            Assert.AreEqual(2, result.Tables[0].Rows[1][0]);
            Assert.AreEqual("How", result.Tables[0].Rows[1][1]);
            Assert.AreEqual(new DateTime(2007, 2, 22), result.Tables[0].Rows[1][2]);

            Assert.AreEqual(3, result.Tables[0].Rows[2][0]);
            Assert.AreEqual("are", result.Tables[0].Rows[2][1]);
            Assert.AreEqual(new DateTime(2002, 1, 19), result.Tables[0].Rows[2][2]);

            Assert.AreEqual("Saturday", result.Tables[0].Rows[3][2]);
            Assert.AreEqual(0.33000000000000002D, result.Tables[0].Rows[4][2]);
            Assert.AreEqual(19, result.Tables[0].Rows[5][2]);
            Assert.AreEqual("Goog", result.Tables[0].Rows[6][2]);
            Assert.AreEqual(12.19D, result.Tables[0].Rows[7][2]);
            Assert.AreEqual(99, result.Tables[0].Rows[8][2]);
            Assert.AreEqual(1385729.234D, result.Tables[0].Rows[9][2]);
        }

        [TestMethod]
        public void GitIssue240ExceptionBeforeRead()
        {
            // Check the exception and message when trying to get data before calling Read().
            // Using the same as SqlDataReader, making it easier to search for a general solution.
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test10x10.xls")))
            {
                var exception = Assert.Throws<InvalidOperationException>(() =>
                {
                    for (int columnIndex = 0; columnIndex < excelReader.FieldCount; columnIndex++)
                    {
                        string _ = excelReader.GetString(columnIndex);
                    }
                });

                Assert.AreEqual("No data exists for the row/column.", exception.Message);
            }
        }

        [TestMethod]
        public void GitIssue241Simple()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_224_simple.xls")))
            {
                Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft åäö &T&CCenter åäö &D&RRight  åäö &P"), "Header");
                Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft åäö &P&CFooter åäö &P&RRight åäö &D"), "Footer");
            }
        }

        [TestMethod]
        public void GitIssue241Simple95()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_224_simple_95.xls")))
            {
                Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft åäö &T&CCenter åäö &D&RRight  åäö &P"), "Header");
                Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft åäö &P&CFooter åäö &P&RRight åäö &D"), "Footer");
            }
        }

        [TestMethod]
        public void GitIssue245CodeName()
        {
            // Test no CodeName = null
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test10x10.xls")))
            {
                Assert.AreEqual(null, reader.CodeName);
            }

            // Test CodeName is set
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Excel_Dataset.xls")))
            {
                Assert.AreEqual("Sheet1", reader.CodeName);
            }

            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_45.xls")))
            {
                Assert.AreEqual("Hoja8", reader.CodeName);
            }
        }

        [TestMethod]
        public void GitIssue250RichText()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_250_richtext.xls")))
            {
                reader.Read();
                var text = reader.GetString(0);
                Assert.AreEqual("Lorem ipsum dolor sit amet, ei pri verterem efficiantur, per id meis idque deterruisset.", text);
            }
        }

        [TestMethod]
        public void GitIssue242Password()
        {
            // BIFF8 standard encryption cryptoapi rc4+sha 
            using (var reader = ExcelReaderFactory.CreateBinaryReader(
                Configuration.GetTestWorkbook("Test_git_issue_242_std_rc4_pwd_password.xls"),
                new ExcelReaderConfiguration { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // Pre-BIFF8 xor obfuscation
            using (var reader = ExcelReaderFactory.CreateBinaryReader(
                Configuration.GetTestWorkbook("Test_git_issue_242_xor_pwd_password.xls"),
                new ExcelReaderConfiguration { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }
        }

        [TestMethod]
        public void BinaryThrowsInvalidPassword()
        {
            Assert.Throws<InvalidPasswordException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateBinaryReader(
                    Configuration.GetTestWorkbook("Test_git_issue_242_xor_pwd_password.xls"),
                    new ExcelReaderConfiguration { Password = "wrongpassword" }))
                {
                    reader.Read();
                }
            });

            Assert.Throws<InvalidPasswordException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateBinaryReader(
                    Configuration.GetTestWorkbook("Test_git_issue_242_xor_pwd_password.xls")))
                {
                    reader.Read();
                }
            });
        }

        [TestMethod]
        public void GitIssue263()
        {
            using (var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test_git_issue_263.xls")))
            {
                var ds = reader.AsDataSet();
                Assert.AreEqual("Economic Inactivity by age\n(Official statistics: not designated as National Statistics)", ds.Tables[1].Rows[3][0]);
            }
        }

        [TestMethod]
        public void BinaryRowHeight()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("CollapsedHide.xls")))
            {
                // Verify the row heights are set when expected, and converted to points from twips
                reader.Read();
                Assert.Greater(reader.RowHeight, 0); 
                Assert.Less(reader.RowHeight, 20);

                reader.Read();
                Assert.Greater(reader.RowHeight, 0);
                Assert.Less(reader.RowHeight, 20);

                reader.Read();
                Assert.Greater(reader.RowHeight, 0);
                Assert.Less(reader.RowHeight, 20);

                reader.Read();
                Assert.AreEqual(0, reader.RowHeight);
            }
        }

        [TestMethod]
        public void GitIssue270EmptyRowsAtTheEnd()
        {
            // AsDataSet() trims trailing blank rows
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_270.xls")))
            {
                var dataSet = reader.AsDataSet();
                Assert.AreEqual(1, dataSet.Tables[0].Rows.Count);
            }

            // Reader methods do not trim trailing blank rows
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_270.xls")))
            {
                var rowCount = 0;
                while (reader.Read())
                    rowCount++;
                Assert.AreEqual(65536, rowCount);
            }
        }

        static bool IsEmptyOrHiddenRow(IExcelDataReader reader)
        {
            if (reader.RowHeight <= 0)
                return true;

            for (var i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetValue(i) != null)
                    return false;
            }

            return true;
        }

        static bool IsEmptyRow(IExcelDataReader reader)
        {
            for (var i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetValue(i) != null)
                    return false;
            }

            return true;
        }

        [TestMethod]
        public void GitIssue160FilterRow()
        {
            // Check there are four rows with data, including empty and hidden rows
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("CollapsedHide.xls")))
            {
                var dataSet = reader.AsDataSet();

                Assert.AreEqual(4, dataSet.Tables[0].Rows.Count);
            }

            // Check there are two rows with content
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("CollapsedHide.xls")))
            {
                var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                    {
                        FilterRow = rowReader => !IsEmptyRow(rowReader)
                    }
                });

                Assert.AreEqual(2, dataSet.Tables[0].Rows.Count);
            }

            // Check there is one visible row with content
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("CollapsedHide.xls")))
            {
                var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                    {
                        FilterRow = rowReader => !IsEmptyOrHiddenRow(rowReader)
                    }
                });

                Assert.AreEqual(1, dataSet.Tables[0].Rows.Count);
            }
        }

        [TestMethod]
        public void GitIssue300FilterColumn()
        {
            // Check there are two columns with data
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("CollapsedHide.xls")))
            {
                var dataSet = reader.AsDataSet();

                Assert.AreEqual(2, dataSet.Tables[0].Columns.Count);
            }

            // Check there is one column when skipping the first
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("CollapsedHide.xls")))
            {
                var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                    {
                        FilterColumn = (rowReader, index) => index > 0
                    }
                });

                Assert.AreEqual(1, dataSet.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void GitIssue265BinaryDisposed()
        {
            var stream = Configuration.GetTestWorkbook("Test10x10.xls");
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream))
            {
                var _ = excelReader.AsDataSet();
            }

            Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
        }

        [TestMethod]
        public void BinaryCompoundLeaveOpen()
        {
            // Verify compound stream is not disposed by the reader
            {
                var stream = Configuration.GetTestWorkbook("Test10x10.xls");
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration()
                {
                    LeaveOpen = true
                }))
                {
                    var _ = excelReader.AsDataSet();
                }

                stream.Seek(0, SeekOrigin.Begin);
                stream.ReadByte();
                stream.Dispose();
            }
        }

        [TestMethod]
        public void BinaryRawLeaveOpen()
        {
            // Verify raw stream is not disposed by the reader
            {
                var stream = Configuration.GetTestWorkbook("biff3.xls");
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration()
                {
                    LeaveOpen = true
                }))
                {
                    var _ = excelReader.AsDataSet();
                }

                stream.Seek(0, SeekOrigin.Begin);
                stream.ReadByte();
                stream.Dispose();
            }
        }

        [TestMethod]
        public void GitIssue286SstStringHeader()
        {
            // Parse xls with SST containing string split exactly between its header and string data across the BIFF Continue records
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_286_SST.xls")))
            {
                Assert.IsNotNull(reader);
            }
        }

        [TestMethod]
        public void GitIssue283TimeSpan()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_283_TimeSpan.xls")))
            {
                reader.Read();
                Assert.AreEqual((TimeSpan)reader[0], new TimeSpan(0));
                Assert.AreEqual((DateTime)reader[2], new DateTime(1899, 12, 31)); // Excel says 1/0/1900, not valid in .NET

                reader.Read();
                Assert.AreEqual((TimeSpan)reader[0], new TimeSpan(1, 0, 0, 0, 0));
                Assert.AreEqual((DateTime)reader[2], new DateTime(1900, 1, 1));

                reader.Read();
                Assert.AreEqual((TimeSpan)reader[0], new TimeSpan(2, 0, 0, 0, 0));

                reader.Read();
                Assert.AreEqual((TimeSpan)reader[0], new TimeSpan(0, 1392, 0, 0, 0));

                reader.Read();
                Assert.AreEqual((TimeSpan)reader[0], new TimeSpan(0, 1416, 0, 0, 0));
                Assert.AreEqual((DateTime)reader[2], new DateTime(1900, 2, 28));

                reader.Read();
                Assert.AreEqual((TimeSpan)reader[0], new TimeSpan(0, 1440, 0, 0, 0));
                Assert.AreEqual((DateTime)reader[2], new DateTime(1900, 2, 28)); // Excel says 2/29/1900, not valid in .NET

                reader.Read();
                Assert.AreEqual((TimeSpan)reader[0], new TimeSpan(0, 1464, 0, 0, 0));
                Assert.AreEqual((DateTime)reader[2], new DateTime(1900, 3, 1));

                reader.Read();
                Assert.AreEqual((TimeSpan)reader[0], new TimeSpan(0, 1488, 0, 0, 0));

                reader.Read();
                Assert.AreEqual((TimeSpan)reader[0], new TimeSpan(0, 1512, 0, 0, 0));
            }
        }


        [TestMethod]
        public void MergedCells()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_MergedCell.xls")))
            {

                excelReader.Read();
                var mergedCells = new List<CellRange> (excelReader.MergeCells);
                Assert.AreEqual(mergedCells.Count, 4, "Incorrect number of merged cells");

                //Sort from top -> left, then down
                mergedCells.Sort(delegate (CellRange c1, CellRange c2)
                {
                    if(c1.FromRow == c2.FromRow)
                    {
                        return c1.FromColumn.CompareTo(c2.FromColumn);
                    }
                    return c1.FromRow.CompareTo(c2.FromRow);
                });

                CollectionAssert.AreEqual(
                    new[]
                    {
                        1,
                        2,
                        0,
                        1
                    },
                    new[]
                    {
                        mergedCells[0].FromRow,
                        mergedCells[0].ToRow,
                        mergedCells[0].FromColumn,
                        mergedCells[0].ToColumn
                    }
                );

                CollectionAssert.AreEqual(
                    new[]
                    {
                        1,
                        5,
                        2,
                        2
                    },
                    new[]
                    {
                        mergedCells[1].FromRow,
                        mergedCells[1].ToRow,
                        mergedCells[1].FromColumn,
                        mergedCells[1].ToColumn
                    }
                );

                CollectionAssert.AreEqual(
                    new[]
                    {
                        3,
                        5,
                        0,
                        0
                    },
                    new[]
                    {
                        mergedCells[2].FromRow,
                        mergedCells[2].ToRow,
                        mergedCells[2].FromColumn,
                        mergedCells[2].ToColumn
                    }
                );

                CollectionAssert.AreEqual(
                    new[]
                    {
                        6,
                        6,
                        0,
                        2
                    },
                    new[]
                    {
                        mergedCells[3].FromRow,
                        mergedCells[3].ToRow,
                        mergedCells[3].FromColumn,
                        mergedCells[3].ToColumn
                    }
                );
            }
        }

        [TestMethod]
        public void GitIssue321MissingEof()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue321.xls")))
            {
                for (int i = 0; i < 7; i++)
                {
                    reader.Read();
                    Assert.IsTrue(string.IsNullOrEmpty(reader.GetString(1)), "Row = " + i);
                }

                reader.Read();
                Assert.AreEqual(" MONETARY AGGREGATES FOR INSTITUTIONAL SECTORS", reader.GetString(1));
            }
        }

        [TestMethod]
        public void GitIssue323DoubleClose()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test10x10.xls")))
            {
                reader.Read();
                reader.Close();
            }
        }

        [TestMethod]
        public void GitIssue329Error()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_329_error.xls")))
            {
                var result = reader.AsDataSet().Tables[0];

                // AsDataSet trims trailing empty rows
                Assert.AreEqual(0, result.Rows.Count);

                // Check errors on first row return null
                reader.Read();
                Assert.IsNull(reader.GetValue(0));
                Assert.IsNull(reader.GetValue(1));
                Assert.IsNull(reader.GetValue(2));
                Assert.AreEqual(1, reader.RowCount);
            }
        }

        [TestMethod]
        public void GitIssue368Header()
        {
            // This reads a specially crafted XLS which loads in Excel:
            // - Raw BIFF5/8 BIFF stream
            // - Non-standard header with size=6, and version=0
            // - Mixes record identifiers from different BIFF versions:
            // - Uses NUMBER (BIFF3-8) and NUMBER_OLD (BIFF2) records
            // - Uses LABEL (BIFF3-5) and LABEL_OLD (BIFF2) records
            // - Uses RK (BIFF3-5) and INTEGER (BIFF2) records
            // - Uses FORMAT_V23 (BIFF2-3) and FORMAT (BIFF4-8) records

            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_368_header.xls")))
            {
                reader.Read();
                Assert.AreEqual("BIFF2", reader[0]);
                Assert.AreEqual(1234.5678, reader[1]);
                Assert.AreEqual(1234, reader[2]);
                Assert.AreEqual("00.0", reader.GetNumberFormatString(1));
                Assert.AreEqual("00.0", reader.GetNumberFormatString(2));

                reader.Read();
                Assert.AreEqual("BIFF3-5", reader[0]);
                Assert.AreEqual(8765.4321, reader[1]);
                Assert.AreEqual(4321, reader[2]);
                Assert.AreEqual("0000.00", reader.GetNumberFormatString(1));
                Assert.AreEqual("0000.00", reader.GetNumberFormatString(2));
            }
        }

        [TestMethod]
        public void GitIssue368Formats()
        {
            // This reads a BIFF2 XLS worksheet created with Excel 2.0 containing 63 number formats, the maximum allowed by the UI.
            // Excel 2.0/2.1 does not write XF/IXFE records, but writes the FORMAT index as 6 bits in the cell attributes.
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_368_formats.xls")))
            {
                for (var i = 0; i < 42; i++)
                {
                    reader.Read();
                    Assert.AreEqual(i % 10, reader[0]);
                    Assert.AreEqual("\"" + i + "\" 0.00", reader.GetNumberFormatString(0));
                }
            }
        }

        [TestMethod]
        public void GitIssue368Ixfe()
        {
            // This reads a specially crafted XLS which loads in Excel:
            // - BIFF2 worksheet, only BIFF2 records
            // - Uses IXFE records to set format
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_368_ixfe.xls")))
            {
                reader.Read();
                Assert.AreEqual("BIFF2", reader[0]);
                Assert.AreEqual(1234.5678, reader[1]);
                Assert.AreEqual(1234, reader[2]);
                Assert.AreEqual("00.0", reader.GetNumberFormatString(1));
                Assert.AreEqual("00.0", reader.GetNumberFormatString(2));

                reader.Read();
                Assert.AreEqual("BIFF2!", reader[0]);
                Assert.AreEqual(8765.4321, reader[1]);
                Assert.AreEqual(4321, reader[2]);
                Assert.AreEqual("0000.00", reader.GetNumberFormatString(1));
                Assert.AreEqual("0000.00", reader.GetNumberFormatString(2));
            }
        }

        [TestMethod]
        public void GitIssue368LabelXf()
        {
            // This reads a specially crafted XLS which loads in Excel:
            // - BIFF2 worksheet, with mixed version FORMAT records, BIFF3-5 label records and 16 bit XF index
            // - Contains 80 XF records
            // - Excel uses only 6 bits of the BIFF3-5 XF index when present in a BIFF2 worksheet, must use IXFE for >62
            // - Excel 2.0 does not write XF>63, but newer Excels read these records
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_368_label_xf.xls")))
            {
                reader.Read();
                Assert.AreEqual("BIFF3-5 record in BIFF2 worksheet with XF 60", reader[0]);
                Assert.AreEqual("\\A@\\B", reader.GetNumberFormatString(0));

                reader.Read();
                Assert.AreEqual("Same with XF 70 (ignored by Excel)", reader[0]);
                // TODO:
                Assert.AreEqual("General", reader.GetNumberFormatString(0));

                reader.Read();
                Assert.AreEqual("Same with XF 70 via IXFE", reader[0]);
                Assert.AreEqual("\\A@\\B", reader.GetNumberFormatString(0));
            }
        }

        [TestMethod]
        public void ColumnWidthsTest()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("ColumnWidthsTest.xls")))
            {
                reader.Read();

                Assert.AreEqual(8.43, reader.GetColumnWidth(0));
                Assert.AreEqual(0, reader.GetColumnWidth(1));
                Assert.AreEqual(15.140625, reader.GetColumnWidth(2));
                Assert.AreEqual(28.7109375, reader.GetColumnWidth(3));

                var expectedException = typeof(ArgumentException);
                var exception = Assert.Throws(expectedException, () =>
                {
                    reader.GetColumnWidth(4);
                });

                Assert.AreEqual($"Column at index 4 does not exist.{Environment.NewLine}Parameter name: i",
                    exception.Message);
            }
        }

        [TestMethod]
        public void GitIssue375IxfeRowMap()
        {
            // This reads a specially crafted XLS which loads in Excel:
            // - 100 rows with IXFE records
            // Verify the internal map of cell offsets used for buffering includes the preceding IXFE records
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_375_ixfe_rowmap.xls")))
            {
                for (var i = 0; i < 100; i++)
                {
                    reader.Read();
                    Assert.AreEqual(1234.0 + i + (i / 10.0), reader[0]);
                    Assert.AreEqual("0.000", reader.GetNumberFormatString(0));
                }
            }
        }

        [TestMethod]
        public void GitIssue382Oom()
        {
            Assert.Throws(typeof(CompoundDocumentException), () =>
            {
                using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_382_oom.xls")))
                {
                    reader.AsDataSet();
                }
            });
        }

        [TestMethod]
        public void GitIssue392Oob()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_392_oob.xls")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(10, result.Rows.Count);
                Assert.AreEqual(10, result.Columns.Count);
                Assert.AreEqual("10x10", result.Rows[1][0]);
                Assert.AreEqual("10x27", result.Rows[9][9]);
            }
        }
    }
}
