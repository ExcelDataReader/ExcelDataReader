using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

using NUnit.Framework;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestCleanup = NUnit.Framework.TearDownAttribute;
using TestInitialize = NUnit.Framework.SetUpAttribute;
using TestMethod = NUnit.Framework.TestAttribute;

namespace ExcelDataReader.Tests
{
    [TestClass]
    public class ExcelOpenXmlReaderTest
    {
        [TestInitialize]
        public void TestInitialize()
        {
#if NETCOREAPP1_0 || NETCOREAPP2_0
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif
        }

        [TestMethod]
        public void GitIssue29ReadSheetStatesReadsCorrectly()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Excel_Dataset.xlsx")))
            {
                Assert.AreEqual("hidden", excelReader.VisibleState);

                excelReader.NextResult();
                Assert.AreEqual("visible", excelReader.VisibleState);

                excelReader.NextResult();
                Assert.AreEqual("veryhidden", excelReader.VisibleState);
            }
        }

        [TestMethod]
        public void GitIssue29AsDataSetProvidesCorrectSheetState()
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Excel_Dataset.xlsx")))
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
        public void Issue11516WorkbookWithSingleSheetShouldNotReturnEmptyDataset()
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11516_Single_Tab.xlsx")))
            {
                Assert.AreEqual(1, reader.ResultsCount);

                DataSet dataSet = reader.AsDataSet();

                Assert.IsTrue(dataSet != null);
                Assert.AreEqual(1, dataSet.Tables.Count);
                Assert.AreEqual(260, dataSet.Tables[0].Rows.Count);
                Assert.AreEqual(29, dataSet.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void AsDataSetTest()
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestOpenXml.xlsx")))
            {
                Assert.AreEqual(3, reader.ResultsCount);

                DataSet dataSet = reader.AsDataSet();

                Assert.IsTrue(dataSet != null);
                Assert.AreEqual(3, dataSet.Tables.Count);
                Assert.AreEqual(7, dataSet.Tables["Sheet1"].Rows.Count);
                Assert.AreEqual(11, dataSet.Tables["Sheet1"].Columns.Count);
            }
        }

        [TestMethod]
        public void ChessTest()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestChess.xlsx")))
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
            using (IExcelDataReader r = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestMultiSheet.xlsx")))
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
                table.TableName = r.Name;

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
                table.TableName = r.Name;
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
            using (IExcelDataReader r = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_num_double_date_bool_string.xlsx")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test10x10000.xlsx")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test10x10.xlsx")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test255x10.xlsx")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestDoublePrecision.xlsx")))
            {
                DataTable result = excelReader.AsDataSet().Tables[0];

                Assert.AreEqual(10, result.Rows.Count);

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
            var expectedException = typeof(Exceptions.HeaderException);

            var exception = Assert.Throws(expectedException, () =>
                {
                    using (ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestFail_Binary.xls")))
                    {
                    }
                });

            Assert.AreEqual("Invalid file signature.", exception.Message);
        }

        [TestMethod]
        public void IssueDateAndTime1468Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Encoding_Formula_Date_1520.xlsx")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                string val1 = new DateTime(2009, 05, 01).ToShortDateString();
                string val2 = DateTime.Parse(dataSet.Tables[0].Rows[1][1].ToString()).ToShortDateString();

                Assert.AreEqual(val1, val2);

                val1 = new DateTime(2009, 1, 1, 11, 0, 0).ToShortTimeString();
                val2 = DateTime.Parse(dataSet.Tables[0].Rows[2][4].ToString()).ToShortTimeString();

                Assert.AreEqual(val1, val2);
            }
        }

        [TestMethod]
        public void Issue8536Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_8536.xlsx")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11397.xlsx")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_4031_NullColumn.xlsx")))
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
        public void Issue4145()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_4145.xlsx")))
            {
                Assert.DoesNotThrow(() => excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration));

                while (excelReader.Read())
                {
                }
            }
        }

        [TestMethod]
        public void Issue10725()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_10725.xlsx")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11435_Colors.xlsx")))
            {
                excelReader.Read();

                Assert.AreEqual("test1", excelReader.GetString(0));
                Assert.AreEqual("test2", excelReader.GetString(1));
                Assert.AreEqual("test3", excelReader.GetString(2));

                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual("test1", dataSet.Tables[0].Rows[0][0].ToString());
                Assert.AreEqual("test2", dataSet.Tables[0].Rows[0][1].ToString());
                Assert.AreEqual("test3", dataSet.Tables[0].Rows[0][2].ToString());
            }
        }

        [TestMethod]
        public void Issue7433IllegalOleAutDate()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_7433_IllegalOleAutDate.xlsx")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_BoolFormula.xlsx")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual(true, dataSet.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void IssueDecimal1109Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Decimal_1109.xlsx")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Encoding_Formula_Date_1520.xlsx")))
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
        public void IssueFileLock5161()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestMultiSheet.xlsx")))
            {
                // read something from the 3rd sheet
                int i = 0;
                do
                {
                    if (i == 0)
                    {
                        excelReader.Read();
                    }
                }
                while (excelReader.NextResult());

                // bug was exposed here
            }
        }

        [TestMethod]
        public void TestBlankHeader()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_BlankHeader.xlsx")))
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
        public void TestOpenOfficeSavedInExcel()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Excel_OpenOffice.xlsx")))
            {
                AssertUtilities.DoOpenOfficeTest(excelReader);
            }
        }

        [TestMethod]
        public void TestIssue11601ReadSheetNames()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Excel_Dataset.xlsx")))
            {
                Assert.AreEqual("test.csv", excelReader.Name);

                excelReader.NextResult();
                Assert.AreEqual("Sheet2", excelReader.Name);

                excelReader.NextResult();
                Assert.AreEqual("Sheet3", excelReader.Name);
            }
        }

        [TestMethod]
        public void MultiSheetTest()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestMultiSheet.xlsx")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_num_double_date_bool_string.xlsx")))
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
                string s = dataSet.Tables[0].Rows[0][2].ToString();
                Assert.AreEqual(new DateTime(2009, 5, 11), DateTime.Parse(s));
                s = dataSet.Tables[0].Rows[29][2].ToString();
                Assert.AreEqual(new DateTime(2009, 11, 30), DateTime.Parse(s));

                // Custom Date Time + Formula
                s = dataSet.Tables[0].Rows[0][5].ToString();
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11479_BlankSheet.xlsx")))
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
        public void Issue11522OpenXml()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11522_OpenXml.xlsx")))
            {
                DataSet result = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(11, result.Tables[0].Columns.Count);
                Assert.AreEqual(1, result.Tables[0].Rows.Count);
                Assert.AreEqual("TestNewButton", result.Tables[0].Rows[0][0]);
                Assert.AreEqual("677", result.Tables[0].Rows[0][1]);
                Assert.AreEqual("u77", result.Tables[0].Rows[0][2]);
                Assert.AreEqual("u766", result.Tables[0].Rows[0][3]);
                Assert.AreEqual("y66", result.Tables[0].Rows[0][4]);
                Assert.AreEqual("F", result.Tables[0].Rows[0][5]);
                Assert.AreEqual(DBNull.Value, result.Tables[0].Rows[0][6]);
                Assert.AreEqual(DBNull.Value, result.Tables[0].Rows[0][7]);
                Assert.AreEqual(DBNull.Value, result.Tables[0].Rows[0][8]);
                Assert.AreEqual(DBNull.Value, result.Tables[0].Rows[0][9]);
                Assert.AreEqual(DBNull.Value, result.Tables[0].Rows[0][10]);
            }
        }

        [TestMethod]
        public void UnicodeCharsTest()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestUnicodeChars.xlsx")))
            {
                DataTable result = excelReader.AsDataSet().Tables[0];

                Assert.AreEqual(3, result.Rows.Count);
                Assert.AreEqual(8, result.Columns.Count);
                Assert.AreEqual("\u00e9\u0417", result.Rows[1][0].ToString());
            }
        }

        /*
        #if !LEGACY
                [TestMethod]
                public void ZipWorker_Extract_Test()
                {
                    var zipper = new ZipWorker(FileSystem.Current, new FileConfiguration.));

                    //this first one isn't a valid xlsx so we are expecting no side effects in the directory tree
                    zipper.Extract(Configuration.GetTestWorkbook("TestChess"));
                    Assert.AreEqual(false, Directory.Exists(zipper.TempPath));
                    Assert.AreEqual(false, zipper.IsValid);

                    //this one is valid so we expect to find the files
                    zipper.Extract(Configuration.GetTestWorkbook("TestOpenXml"));

                    Assert.AreEqual(true, Directory.Exists(zipper.TempPath));
                    Assert.AreEqual(true, zipper.IsValid);

                    string tPath = zipper.TempPath;

                    //make sure that dispose gets rid of the files
                    zipper.Dispose();

                    Assert.AreEqual(false, Directory.Exists(tPath));
                }

                private class FileConfiguration.: IFileConfiguration.
                {
                    public string GetTempPath()
                    {
                        return System.IO.Path.GetTempPath();
                    }
                }
        #endif
        */

        [TestMethod]
        public void IssueDateFormatButNotDate()
        {
            // we want to make sure that if a cell is formatted as a date but it's contents are not a date then
            // the output is not a date (it was ending up as datetime.min)
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_DateFormatButNotDate.xlsx")))
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
        public void Issue11573BlankValues()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11573_BlankValues.xlsx")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual(1D, dataSet.Tables[0].Rows[12][0]);
                Assert.AreEqual("070202", dataSet.Tables[0].Rows[12][1]);
            }
        }

        [TestMethod]
        public void Issue11773Exponential()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11773_Exponential.xlsx")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(2566.37168141593D, double.Parse(dataSet.Tables[0].Rows[0][6].ToString()));
            }
        }

        [TestMethod]
        public void Issue11773ExponentialCommas()
        {
#if NETCOREAPP1_0
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");
#else
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE", false);
#endif

            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11773_Exponential.xlsx")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(2566.37168141593D, double.Parse(dataSet.Tables[0].Rows[0][6].ToString()));
            }
        }

        [TestMethod]
        public void TestGoogleSourced()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_googlesourced.xlsx")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual("9583638582", dataSet.Tables[0].Rows[0][0].ToString());
                Assert.AreEqual(4, dataSet.Tables[0].Rows.Count);
                Assert.AreEqual(6, dataSet.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void TestIssue12667GoogleExportMissingColumns()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_12667_GoogleExport_MissingColumns.xlsx")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(7, dataSet.Tables[0].Columns.Count); // 6 with data + 1 that is present but no data in it
                Assert.AreEqual(0, dataSet.Tables[0].Rows.Count);
            }
        }

        /// <summary>
        /// Makes sure that we can read data from the first row of last sheet
        /// </summary>
        [TestMethod]
        public void Issue12271NextResultSet()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_LotsOfSheets.xlsx")))
            {
                do
                {
                    excelReader.Read();

                    if (excelReader.FieldCount == 0)
                    {
                        continue;
                    }

                    // ignore sheets beginning with $e
                    if (excelReader.Name.StartsWith("$e", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    Assert.AreEqual("StaffName", excelReader.GetString(0));
                }
                while (excelReader.NextResult());
            }
        }

        [TestMethod]
        public void IssueGit142()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_Issue_142.xlsx")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual(4, dataSet.Tables[0].Columns.Count);
            }
        }

        /// <summary>
        /// Sheet has no [dimension] and/or no [cols].
        /// Sheet has no [styles].
        /// Each row [row] has no "r" attribute.
        /// Each cell [c] has no "r" attribute.
        /// </summary>
        [TestMethod]
        public void IssueNoStylesNoRAttribute()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_NoStyles_NoRAttribute.xlsx")))
            {
                DataSet result = excelReader.AsDataSet();

                Assert.IsTrue(result.Tables.Count > 0);
                Assert.AreEqual(39, result.Tables[0].Rows.Count);
                Assert.AreEqual(18, result.Tables[0].Columns.Count);
                Assert.AreEqual("ROW NUMBER 5", result.Tables[0].Rows[4][4].ToString());

                excelReader.Close();
            }
        }

        [TestMethod]
        public void NoDimensionOrCellReferenceAttribute()
        {
            // 20170306_Daily Package GPR 250 Index EUR Overview.xlsx
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("NoDimensionOrCellReferenceAttribute.xlsx")))
            {
                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(2, result.Tables.Count);
                Assert.AreEqual(8, result.Tables[0].Columns.Count, "Sheet0 Columns");
                Assert.AreEqual(7, result.Tables[0].Rows.Count, "Sheet0 Rows");

                Assert.AreEqual(8, result.Tables[1].Columns.Count, "Sheet1 Columns");
                Assert.AreEqual(20, result.Tables[1].Rows.Count, "Sheet1 Rows");
            }
        }

        [TestMethod]
        public void CellValueIso8601Date()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_221.xlsx")))
            {
                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(new DateTime(2017, 3, 16), result.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void CellFormat49()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Format49_@.xlsx")))
            {
                DataSet result = excelReader.AsDataSet();

                // ExcelDataReader used to convert numbers formatted with NumFmtId=49/@ to culture-specific strings.
                // This behaviour changed in v3 to return the original value:
                // Assert.That(result.Tables[0].Rows[0].ItemArray, Is.EqualTo(new[] { "2010-05-05", "1.1", "2,2", "123", "2,2" }));
                Assert.That(result.Tables[0].Rows[0].ItemArray, Is.EqualTo(new object[] { "2010-05-05", "1.1", 2.2000000000000002D, 123.0D, "2,2" }));
            }
        }

        [TestMethod]
        public void GitIssue97()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("fillreport.xlsx")))
            {
                // fillreport.xlsx was generated by a third party and uses badly formatted cell references with only numerals.
                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(1, result.Tables.Count);
                Assert.AreEqual(20, result.Tables[0].Rows.Count);
                Assert.AreEqual(10, result.Tables[0].Columns.Count);
                Assert.AreEqual("Account Number", result.Tables[0].Rows[1][0]);
                Assert.AreEqual("Trader", result.Tables[0].Rows[1][1]);
            }
        }

        [TestMethod]
        public void GitIssue82Date1900OpenXml()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("roo_1900_base.xlsx")))
            {
                // 15/06/2009
                // 4/19/2013 (=TODAY() when file was saved)

                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(new DateTime(2009, 6, 15), (DateTime)result.Tables[0].Rows[0][0]);
                Assert.AreEqual(new DateTime(2013, 4, 19), (DateTime)result.Tables[0].Rows[1][0]);
            }
        }

        [TestMethod]
        //        [Ignore("Pending fix")]
        public void GitIssue82Date1904OpenXml()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("roo_1904_base.xlsx")))
            {
                // 15/06/2009
                // 4/19/2013 (=TODAY() when file was saved)

                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(new DateTime(2009, 6, 15), (DateTime)result.Tables[0].Rows[0][0]);
                Assert.AreEqual(new DateTime(2013, 4, 19), (DateTime)result.Tables[0].Rows[1][0]);
            }
        }

        [TestMethod]
        public void GitIssue68NullSheetPath()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_68_NullSheetPath.xlsm")))
            {
                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(2, result.Tables[0].Columns.Count);
                Assert.AreEqual(1, result.Tables[0].Rows.Count);

            }
        }

        [TestMethod]
        public void GitIssue53CachedFormulaStringType()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_53_Cached_Formula_String_Type.xlsx")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                // Ensure that parseable, numeric cached formula values are read as a double
                Assert.IsInstanceOf<double>(dataSet.Tables[0].Rows[0][2]);
                Assert.AreEqual(3D, dataSet.Tables[0].Rows[0][2]);

                // Ensure that non-parseable, non-numeric cached formula values are read as a string
                Assert.IsInstanceOf<string>(dataSet.Tables[0].Rows[1][2]);
                Assert.AreEqual("AB", dataSet.Tables[0].Rows[1][2]);

                // Ensure that parseable, non-numeric cached formula values are read as a string
                Assert.IsInstanceOf<string>(dataSet.Tables[0].Rows[2][2]);
                Assert.AreEqual("1,", dataSet.Tables[0].Rows[2][2]);
            }
        }

        [TestMethod]
        public void GitIssue14InvalidOADate()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_14_InvalidOADate.xlsx")))
            {
                var dataSet = excelReader.AsDataSet();

                // Test out of range double formatted as date returns double
                Assert.AreEqual(1000000000000D, dataSet.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void GitIssue241Simple()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_224_simple.xlsx")))
            {
                Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft åäö &T&CCenter åäö &D&RRight  åäö &P"), "Header");
                Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft åäö &P&CFooter åäö &P&RRight åäö &D"), "Footer");
            }
        }

        [TestMethod]
        public void GitIssue241FirstOddEven()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_224_firstoddeven.xlsx")))
            {
                Assert.That(reader.HeaderFooter, Is.Not.Null);

                Assert.That(reader.HeaderFooter?.HasDifferentFirst, Is.True, "HasDifferentFirst");
                Assert.That(reader.HeaderFooter?.HasDifferentOddEven, Is.True, "HasDifferentOddEven");

                Assert.That(reader.HeaderFooter?.FirstHeader, Is.EqualTo("&CFirst header center"), "First Header");
                Assert.That(reader.HeaderFooter?.FirstFooter, Is.EqualTo("&CFirst footer center"), "First Footer");
                Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft åäö &T&COdd page header&RRight  åäö &P"), "Odd Header");
                Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft åäö &P&COdd Footer åäö &P&RRight åäö &D"), "Odd Footer");
                Assert.That(reader.HeaderFooter?.EvenHeader, Is.EqualTo("&L&A&CEven page header"), "Even Header");
                Assert.That(reader.HeaderFooter?.EvenFooter, Is.EqualTo("&CEven page footer"), "Even Footer");
            }
        }

        [TestMethod]
        public void GitIssue245CodeName()
        {
            // Test no codename = null
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test10x10.xlsx")))
            {
                Assert.AreEqual(null, reader.CodeName);
            }

            // Test CodeName is set
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Excel_Dataset.xlsx")))
            {
                Assert.AreEqual("Sheet1", reader.CodeName);
            }
        }

        [TestMethod]
        public void GitIssue250RichText()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_250_richtext.xlsx")))
            {
                reader.Read();
                var text = reader.GetString(0);
                Assert.AreEqual("Lorem ipsum dolor sit amet, ei pri verterem efficiantur, per id meis idque deterruisset.", text);
            }
        }

        [TestMethod]
        public void GitIssue242StandardEncryption()
        {
            // OpenXml standard encryption aes128+sha1
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("standard_AES128_SHA1_ECB_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml standard encryption aes192+sha1
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("standard_AES192_SHA1_ECB_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml standard encryption aes256+sha1
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("standard_AES256_SHA1_ECB_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }
        }

        [TestMethod]
        public void GitIssue242AgileEncryption()
        {
            // OpenXml agile encryption aes128+md5+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES128_MD5_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes128+sha1+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES128_SHA1_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes128+sha384+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES128_SHA384_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes128+sha512+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES128_SHA512_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes192+sha512+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES192_SHA512_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes256+sha512+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES256_SHA512_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption 3des+sha384+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_DESede_SHA384_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // The following encryptions do not exist on netstandard1.3
#if NET20 || NET45 || NETCOREAPP2_0
            // OpenXml agile encryption des+md5+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_DES_MD5_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption rc2+sha1+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_RC2_SHA1_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }
#endif
        }

        [TestMethod]
        public void OpenXmlThrowsInvalidPassword()
        {
            Assert.Throws<Exceptions.InvalidPasswordException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                    Configuration.GetTestWorkbook("agile_AES128_MD5_CBC_pwd_password.xlsx"),
                    new ExcelReaderConfiguration() { Password = "wrongpassword" }))
                {
                    reader.Read();
                }
            });

            Assert.Throws<Exceptions.InvalidPasswordException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                    Configuration.GetTestWorkbook("agile_AES128_MD5_CBC_pwd_password.xlsx")))
                {
                    reader.Read();
                }
            });
        }

        [TestMethod]
        public void OpenXmlThrowsEmptyZipFile()
        {
            Assert.Throws<Exceptions.HeaderException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("EmptyZipFile.xlsx")))
                {
                    reader.Read();
                }
            });
        }

        [TestMethod]
        public void OpenXmlRowHeight()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("CollapsedHide.xlsx")))
            {
                reader.Read();
                Assert.Greater(reader.RowHeight, 0);

                reader.Read();
                Assert.Greater(reader.RowHeight, 0);

                reader.Read();
                Assert.Greater(reader.RowHeight, 0);

                reader.Read();
                Assert.AreEqual(reader.RowHeight, 0);
            }
        }

        [TestMethod]
        public void GitIssue270EmptyRowsAtTheEnd()
        {
            // AsDataSet() trims trailing blank rows
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_270.xlsx")))
            {
                var dataSet = reader.AsDataSet();
                Assert.AreEqual(1, dataSet.Tables[0].Rows.Count);
            }

            // Reader methods do not trim trailing blank rows
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_270.xlsx")))
            {
                var rowCount = 0;
                while (reader.Read())
                    rowCount++;
                Assert.AreEqual(65536, rowCount);
            }
        }

        [TestMethod]
        public void GitIssue265OpenXmlDisposed()
        {
            // Verify the file stream is closed and disposed by the reader
            {
                var stream = Configuration.GetTestWorkbook("Test10x10.xlsx");
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    var _ = excelReader.AsDataSet();
                }

                Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
            }

            // Verify streams used by standard encryption are closed
            {
                var stream = Configuration.GetTestWorkbook("standard_AES128_SHA1_ECB_pwd_password.xlsx");

                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(
                    stream,
                    new ExcelReaderConfiguration() { Password = "password" }))
                {
                    var _ = excelReader.AsDataSet();
                }

                Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
            }

            // Verify streams used by agile encryption are closed
            {
                var stream = Configuration.GetTestWorkbook("agile_AES128_MD5_CBC_pwd_password.xlsx");

                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(
                    stream,
                    new ExcelReaderConfiguration() { Password = "password" }))
                {
                    var _ = excelReader.AsDataSet();
                }

                Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
            }
        }

        [TestMethod]
        public void OpenXmlLeaveOpen()
        {
            // Verify the file stream is closed and disposed by the reader
            {
                var stream = Configuration.GetTestWorkbook("Test10x10.xlsx");
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream, new ExcelReaderConfiguration()
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
        public void GitIssue271InvalidDimension()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_271_InvalidDimension.xlsx")))
            {
                var dataSet = excelReader.AsDataSet();
                Assert.AreEqual(3, dataSet.Tables[0].Columns.Count);
                Assert.AreEqual(9, dataSet.Tables[0].Rows.Count);
            }
        }

        [TestMethod]
        public void GitIssue283TimeSpan()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_283_TimeSpan.xlsx")))
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
        public void GitIssue289CompoundDocumentEncryptedWithDefaultPassword()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue289.xlsx")))
            {
                reader.Read();
                Assert.AreEqual("aaaaaaa", reader.GetValue(0));
            }
        }

        [TestMethod]
        public void MergedCells()
        {
            // XLSX was manually edited to include a <mergecell></mergecell> element with closing tag
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_MergedCell.xlsx")))
            {
                excelReader.Read();
                var mergedCells = new List<CellRange>(excelReader.MergeCells);
                Assert.AreEqual(mergedCells.Count, 4, "Incorrect number of merged cells");

                //Sort from top -> left, then down
                mergedCells.Sort(delegate (CellRange c1, CellRange c2)
                {
                    if (c1.FromRow == c2.FromRow)
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
        public void GitIssue301IgnoreCase()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_301_IgnoreCase.xlsx")))
            {
                DataTable result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(10, result.Rows.Count);
                Assert.AreEqual(10, result.Columns.Count);
                Assert.AreEqual("10x10", result.Rows[1][0]);
                Assert.AreEqual("10x27", result.Rows[9][9]);
            }
        }

        [TestMethod]
        public void GitIssue319InlineRichText()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue319.xlsx")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual("Text1", result.Rows[0][0]);
            }
        }

        [TestMethod]
        public void GitIssue324MultipleRowElementsPerRow()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_324.xlsx")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(20, result.Rows.Count);
                Assert.AreEqual(13, result.Columns.Count);

                Assert.That(result.Rows[10].ItemArray, Is.EqualTo(new object[] { DBNull.Value, DBNull.Value, "Other", 191036.15, 194489.45, 66106.32, 37167.88, 102589.54, 57467.94, 130721.93, 150752.67, 76300.69, 67024.6 }));
            }
        }

        [TestMethod]
        public void GitIssue323DoubleClose()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test10x10.xlsx")))
            {
                reader.Read();
                reader.Close();
            }
        }

        [TestMethod]
        public void GitIssue329Error()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_329_error.xlsx")))
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
        public void GitIssue354()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("test_git_issue_354.xlsx")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(1, result.Rows.Count);
                Assert.AreEqual("cell data", result.Rows[0][0]);
            }
        }

        [TestMethod]
        public void GitIssue364()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("test_git_issue_364.xlsx")))
            {
                Assert.AreEqual(1, reader.RowCount);
                reader.Read();

                Assert.AreEqual(0, reader.GetNumberFormatIndex(0));
                Assert.AreEqual(-1, reader.GetNumberFormatIndex(1));
                Assert.AreEqual(14, reader.GetNumberFormatIndex(2));
                Assert.AreEqual(164, reader.GetNumberFormatIndex(3));
            }
        }

        [TestMethod]
        public void ColumnWidthsTest()
        {
            // XLSX was manually edited to include a <col></col> element with closing tag
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("ColumnWidthsTest.xlsx")))
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
        public void GitIssue385Backslash()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_385_backslash.xlsx")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(10, result.Rows.Count);
                Assert.AreEqual(10, result.Columns.Count);
                Assert.AreEqual("10x10", result.Rows[1][0]);
                Assert.AreEqual("10x27", result.Rows[9][9]);
            }
        }

        [TestMethod]
        public void GitIssue_341_Indent()
        {
            int[][] expected =
            {
                new[] { 2, 0, 0 },
                new[] { 2, 0, 0 },
                new[] { 3, 3, 4 },
                new[] { 1, 1, 1 }, // Merged cell
                new[] { 2, 0, 0 },
            };

            int index = 0;
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_341.xlsx")))
            {
                while (reader.Read())
                {
                    int[] expectedRow = expected[index];
                    int[] actualRow = new int[reader.FieldCount];
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        actualRow[i] = reader.GetCellStyle(i).IndentLevel;
                    }

                    Assert.AreEqual(expectedRow, actualRow, "Indent level on row '{0}'.", index);

                    index++;
                }
            }
        }

        [TestMethod]
        public void GitIssue_341_HorizontalAlignment()
        {
            HorizontalAlignment[][] expected =
            {
                new[] { HorizontalAlignment.Left, HorizontalAlignment.General, HorizontalAlignment.General },
                new[] { HorizontalAlignment.Distributed, HorizontalAlignment.General, HorizontalAlignment.General },
                new[] { HorizontalAlignment.Left, HorizontalAlignment.Left, HorizontalAlignment.Left },
                new[] { HorizontalAlignment.Left, HorizontalAlignment.Left, HorizontalAlignment.Left }, // Merged cell
                new[] { HorizontalAlignment.Left, HorizontalAlignment.General, HorizontalAlignment.General },
            };

            int index = 0;
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_341.xlsx")))
            {
                while (reader.Read())
                {
                    HorizontalAlignment[] expectedRow = expected[index];
                    HorizontalAlignment[] actualRow = new HorizontalAlignment[reader.FieldCount];
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        actualRow[i] = reader.GetCellStyle(i).HorizontalAlignment;
                    }

                    Assert.AreEqual(expectedRow, actualRow, "Horizontal alignment on row '{0}'.", index);

                    index++;
                }
            }
        }
    }
}
