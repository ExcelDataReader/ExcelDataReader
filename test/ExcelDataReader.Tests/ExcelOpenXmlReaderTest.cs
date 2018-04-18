using System;
using System.Collections.Generic;
using System.Globalization;
#if NET20 || NET45 || NETCOREAPP2_0
using System.Data;
#endif
using ExcelDataReader.Tests;

using NUnit.Framework;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestCleanup = NUnit.Framework.TearDownAttribute;
using TestInitialize = NUnit.Framework.SetUpAttribute;
using TestMethod = NUnit.Framework.TestAttribute;

#if EXCELDATAREADER_NET20
namespace ExcelDataReader.Net20.Tests
#elif NET45
namespace ExcelDataReader.Net45.Tests
#elif NETCOREAPP1_0
namespace ExcelDataReader.Netstandard13.Tests
#elif NETCOREAPP2_0
namespace ExcelDataReader.Netstandard20.Tests
#else
#error "Tests do not support the selected target platform"
#endif
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
        public void GitIssue_29_ReadSheetStatesReadsCorrectly()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Excel_Dataset")))
            {
                Assert.AreEqual("hidden", excelReader.VisibleState);

                excelReader.NextResult();
                Assert.AreEqual("visible", excelReader.VisibleState);

                excelReader.NextResult();
                Assert.AreEqual("veryhidden", excelReader.VisibleState);
            }
        }

        [TestMethod]
        public void GitIssue_29_AsDataSetProvidesCorrectSheetState()
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Excel_Dataset")))
            {
                var dataset = reader.AsDataSet();

                Assert.IsTrue(dataset != null);
                Assert.AreEqual(3, dataset.Tables.Count);
                Assert.AreEqual("hidden", dataset.Tables[0].ExtendedProperties["visiblestate"]);
                Assert.AreEqual("visible", dataset.Tables[1].ExtendedProperties["visiblestate"]);
                Assert.AreEqual("veryhidden", dataset.Tables[2].ExtendedProperties["visiblestate"]);
            }
        }

        [TestMethod]
        public void Issue_11516_workbook_with_single_sheet_should_not_return_empty_dataset()
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_11516_Single_Tab")))
            {
                Assert.AreEqual(1, reader.ResultsCount);

                DataSet dataset = reader.AsDataSet();

                Assert.IsTrue(dataset != null);
                Assert.AreEqual(1, dataset.Tables.Count);
                Assert.AreEqual(260, dataset.Tables[0].Rows.Count);
                Assert.AreEqual(29, dataset.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void AsDataset_Test()
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTestOpenXml")))
            {
                Assert.AreEqual(3, reader.ResultsCount);

                DataSet dataset = reader.AsDataSet();

                Assert.IsTrue(dataset != null);
                Assert.AreEqual(3, dataset.Tables.Count);
                Assert.AreEqual(7, dataset.Tables["Sheet1"].Rows.Count);
                Assert.AreEqual(11, dataset.Tables["Sheet1"].Columns.Count);
            }
        }

        [TestMethod]
        public void ChessTest()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTestChess")))
            {
                DataTable result = excelReader.AsDataSet().Tables[0];

                Assert.AreEqual(4, result.Rows.Count);
                Assert.AreEqual(6, result.Columns.Count);
                Assert.AreEqual("1", result.Rows[3][5].ToString());
                Assert.AreEqual("1", result.Rows[2][0].ToString());
            }
        }

        [TestMethod]
        public void DataReader_NextResult_Test()
        {
            using (IExcelDataReader r = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTestMultiSheet")))
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
        public void DataReader_Read_Test()
        {
            using (IExcelDataReader r = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_num_double_date_bool_string")))
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
        public void Dimension10x10000Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest10x10000")))
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
        public void Dimension10x10Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest10x10")))
            {
                DataTable result = excelReader.AsDataSet().Tables[0];

                Assert.AreEqual(10, result.Rows.Count);
                Assert.AreEqual(10, result.Columns.Count);
                Assert.AreEqual("10x10", result.Rows[1][0]);
                Assert.AreEqual("10x27", result.Rows[9][9]);
            }
        }

        [TestMethod]
        public void Dimension255x10Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest255x10")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTestDoublePrecision")))
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
        public void Fail_Test()
        {
            var expectedException = typeof(Exceptions.HeaderException);

            var exception = Assert.Throws(expectedException, () =>
                {
                    using (ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestFail_Binary")))
                    {
                    }
                });

            Assert.AreEqual("Invalid file signature.", exception.Message);
        }

        [TestMethod]
        public void Issue_Date_and_Time_1468_Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Encoding_1520")))
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
        public void Issue_8536_Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_8536")))
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
                var datenum1 = dataSet.Tables[0].Rows[5][1];
                Assert.AreEqual(typeof(double), datenum1.GetType());
                Assert.AreEqual(41244, double.Parse(datenum1.ToString()));

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
        public void Issue_11397_Currency_Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_11397")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual("$44.99", dataSet.Tables[0].Rows[1][0].ToString()); // general in spreadsheet so should be a string including the $
                Assert.AreEqual(44.99, double.Parse(dataSet.Tables[0].Rows[2][0].ToString())); // currency euros in spreadsheet so should be a currency
                Assert.AreEqual(44.99, double.Parse(dataSet.Tables[0].Rows[3][0].ToString())); // currency pounds in spreadsheet so should be a currency
            }
        }

        [TestMethod]
        public void Issue_4031_NullColumn()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_4031_NullColumn")))
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
        public void Issue_4145()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_4145")))
            {
                Assert.DoesNotThrow(() => excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration));

                while (excelReader.Read())
                {
                }
            }
        }

        [TestMethod]
        public void Issue_10725()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_10725")))
            {
                excelReader.Read();
                Assert.AreEqual(8.8, excelReader.GetValue(0));

                DataSet result = excelReader.AsDataSet();

                Assert.AreEqual(8.8, result.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void Issue_11435_Colors()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_11435_Colors")))
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
        public void Issue_7433_IllegalOleAutDate()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_7433_IllegalOleAutDate")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual(3.10101195608231E+17, dataSet.Tables[0].Rows[0][0]);
                Assert.AreEqual("B221055625", dataSet.Tables[0].Rows[1][0]);
                Assert.AreEqual(4.12721197309241E+17, dataSet.Tables[0].Rows[2][0]);
            }
        }

        [TestMethod]
        public void Issue_BoolFormula()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_BoolFormula")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual(true, dataSet.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void Issue_Decimal_1109_Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Decimal_1109")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual(3.14159, dataSet.Tables[0].Rows[0][0]);

                const double val1 = -7080.61;
                double val2 = (double)dataSet.Tables[0].Rows[0][1];
                Assert.AreEqual(val1, val2);
            }
        }

        [TestMethod]
        public void Issue_Encoding_1520_Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Encoding_1520")))
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
        public void Issue_FileLock_5161()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTestMultiSheet")))
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
        public void Test_BlankHeader()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_BlankHeader")))
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
        public void Test_OpenOffice_SavedInExcel()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Excel_OpenOffice")))
            {
                AssertUtilities.DoOpenOfficeTest(excelReader);
            }
        }

        [TestMethod]
        public void Test_Issue_11601_ReadSheetnames()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Excel_Dataset")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTestMultiSheet")))
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
        public void Test_num_double_date_bool_string()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_num_double_date_bool_string")))
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
        public void Issue_11479_BlankSheet()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xIssue_11479_BlankSheet")))
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
        public void Issue_11522_OpenXml()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_11522_OpenXml")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTestUnicodeChars")))
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
            zipper.Extract(Configuration.GetTestWorkbook("xTestOpenXml"));

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
        public void Issue_DateFormatButNotDate()
        {
            // we want to make sure that if a cell is formatted as a date but it's contents are not a date then
            // the output is not a date (it was ending up as datetime.min)
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_DateFormatButNotDate")))
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
        public void Issue_11573_BlankValues()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_11573_BlankValues")))
            {
                var dataset = excelReader.AsDataSet();

                Assert.AreEqual(1D, dataset.Tables[0].Rows[12][0]);
                Assert.AreEqual("070202", dataset.Tables[0].Rows[12][1]);
            }
        }

        [TestMethod]
        public void Issue_11773_Exponential()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_11773_Exponential")))
            {
                var dataset = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(2566.37168141593D, double.Parse(dataset.Tables[0].Rows[0][6].ToString()));
            }
        }

        [TestMethod]
        public void Issue_11773_Exponential_Commas()
        {
#if NETCOREAPP1_0
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");
#else
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE", false);
#endif

            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_11773_Exponential")))
            {
                var dataset = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(2566.37168141593D, double.Parse(dataset.Tables[0].Rows[0][6].ToString()));
            }
        }

        [TestMethod]
        public void Test_googlesourced()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_googlesourced")))
            {
                var dataset = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual("9583638582", dataset.Tables[0].Rows[0][0].ToString());
                Assert.AreEqual(4, dataset.Tables[0].Rows.Count);
                Assert.AreEqual(6, dataset.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void Test_Issue_12667_GoogleExport_MissingColumns()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Issue_12667_GoogleExport_MissingColumns")))
            {
                var dataset = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(7, dataset.Tables[0].Columns.Count); // 6 with data + 1 that is present but no data in it
                Assert.AreEqual(0, dataset.Tables[0].Rows.Count);
            }
        }

        /// <summary>
        /// Makes sure that we can read data from the first roiw of last sheet
        /// </summary>
        [TestMethod]
        public void Issue_12271_NextResultSet()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_LotsOfSheets")))
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
        public void Issue_Git_142()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_Issue_142")))
            {
                var dataset = excelReader.AsDataSet();

                Assert.AreEqual(4, dataset.Tables[0].Columns.Count);
            }
        }

        /// <summary>
        /// Sheet has no [dimension] and/or no [cols].
        /// Sheet has no [styles].
        /// Each row [row] has no "r" attribute.
        /// Each cell [c] has no "r" attribute.
        /// </summary>
        [TestMethod]
        public void Issue_NoStyles_NoRAttribute()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_NoStyles_NoRAttribute")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("NoDimensionOrCellReferenceAttribute")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_221")))
            {
                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(new DateTime(2017, 3, 16), result.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void CellFormat49()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Format49_@")))
            {
                DataSet result = excelReader.AsDataSet();

                // ExcelDataReader used to convert numbers formatted with NumFmtId=49/@ to culture-specific strings.
                // This behaviour changed in v3 to return the original value:
                // Assert.That(result.Tables[0].Rows[0].ItemArray, Is.EqualTo(new[] { "2010-05-05", "1.1", "2,2", "123", "2,2" }));
                Assert.That(result.Tables[0].Rows[0].ItemArray, Is.EqualTo(new object[] { "2010-05-05", "1.1", 2.2000000000000002D, 123.0D, "2,2" }));
            }
        }

        [TestMethod]
        public void GitIssue_97()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("fillreport")))
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
        public void GitIssue_82_Date1900_OpenXml()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xroo_1900_base")))
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
        public void GitIssue_82_Date1904_OpenXml()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xroo_1904_base")))
            {
                // 15/06/2009
                // 4/19/2013 (=TODAY() when file was saved)

                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(new DateTime(2009, 6, 15), (DateTime)result.Tables[0].Rows[0][0]);
                Assert.AreEqual(new DateTime(2013, 4, 19), (DateTime)result.Tables[0].Rows[1][0]);
            }
        }

        [TestMethod]
        public void GitIssue_68_NullSheetPath()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_68_NullSheetPath")))
            {
                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(2, result.Tables[0].Columns.Count);
                Assert.AreEqual(1, result.Tables[0].Rows.Count);

            }
        }

        [TestMethod]
        public void GitIssue_53_Cached_Formula_String_Type()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_53_Cached_Formula_String_Type")))
            {
                var dataset = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                // Ensure that parseable, numeric cached formula values are read as a double
                Assert.IsInstanceOf<double>(dataset.Tables[0].Rows[0][2]);
                Assert.AreEqual(3D, dataset.Tables[0].Rows[0][2]);

                // Ensure that non-parseable, non-numeric cached formula values are read as a string
                Assert.IsInstanceOf<string>(dataset.Tables[0].Rows[1][2]);
                Assert.AreEqual("AB", dataset.Tables[0].Rows[1][2]);

                // Ensure that parseable, non-numeric cached formula values are read as a string
                Assert.IsInstanceOf<string>(dataset.Tables[0].Rows[2][2]);
                Assert.AreEqual("1,", dataset.Tables[0].Rows[2][2]);
            }
        }

        [TestMethod]
        public void GitIssue_14_InvalidOADate()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_14_InvalidOADate")))
            {
                var dataset = excelReader.AsDataSet();

                // Test out of range double formatted as date returns double
                Assert.AreEqual(1000000000000D, dataset.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void GitIssue_241_Simple()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_224_simple")))
            {
                Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft åäö &T&CCenter åäö &D&RRight  åäö &P"), "Header");
                Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft åäö &P&CFooter åäö &P&RRight åäö &D"), "Footer");
            }
        }

        [TestMethod]
        public void GitIssue_241_FirstOddEven()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_224_firstoddeven")))
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
        public void GitIssue_245_CodeName()
        {
            // Test no codename = null
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest10x10")))
            {
                Assert.AreEqual(null, reader.CodeName);
            }

            // Test CodeName is set
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_Excel_Dataset")))
            {
                Assert.AreEqual("Sheet1", reader.CodeName);
            }
        }

        [TestMethod]
        public void GitIssue_250_RichText()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_git_issue_250_richtext")))
            {
                reader.Read();
                var text = reader.GetString(0);
                Assert.AreEqual("Lorem ipsum dolor sit amet, ei pri verterem efficiantur, per id meis idque deterruisset.", text);
            }
        }

        [TestMethod]
        public void GitIssue_242_StandardEncryption()
        {
            // OpenXml standard encryption aes128+sha1
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("standard_AES128_SHA1_ECB_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml standard encryption aes192+sha1
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("standard_AES192_SHA1_ECB_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml standard encryption aes256+sha1
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("standard_AES256_SHA1_ECB_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }
        }

        [TestMethod]
        public void GitIssue_242_AgileEncryption()
        {
            // OpenXml agile encryption aes128+md5+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES128_MD5_CBC_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes128+sha1+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES128_SHA1_CBC_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes128+sha384+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES128_SHA384_CBC_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes128+sha512+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES128_SHA512_CBC_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes192+sha512+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES192_SHA512_CBC_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes256+sha512+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES256_SHA512_CBC_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption 3des+sha384+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_DESede_SHA384_CBC_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // The following encryptions do not exist on netstandard1.3
#if NET20 || NET45 || NETCOREAPP2_0
            // OpenXml agile encryption des+md5+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_DES_MD5_CBC_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption rc2+sha1+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_RC2_SHA1_CBC_pwd_password"),
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
                    Configuration.GetTestWorkbook("agile_AES128_MD5_CBC_pwd_password"),
                    new ExcelReaderConfiguration() { Password = "wrongpassword" }))
                {
                    reader.Read();
                }
            });

            Assert.Throws<Exceptions.InvalidPasswordException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                    Configuration.GetTestWorkbook("agile_AES128_MD5_CBC_pwd_password")))
                {
                    reader.Read();
                }
            });
        }

        [TestMethod]
        public void OpenXmlThrowsEmptyZipfile()
        {
            Assert.Throws<Exceptions.HeaderException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("EmptyZipFile")))
                {
                    reader.Read();
                }
            });
        }

        [TestMethod]
        public void OpenXmlRowHeight()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xCollapsedHide")))
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
        public void GitIssue_270_EmptyRowsAtTheEnd()
        {
            // AsDataSet() trims trailing blank rows
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_git_issue_270")))
            {
                var dataset = reader.AsDataSet();
                Assert.AreEqual(1, dataset.Tables[0].Rows.Count);
            }

            // Reader methods do not trim trailing blank rows
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_git_issue_270")))
            {
                var rowCount = 0;
                while (reader.Read())
                    rowCount++;
                Assert.AreEqual(65536, rowCount);
            }
        }

        [TestMethod]
        public void GitIssue_265_OpenXmlDisposed()
        {
            // Verify the file stream is closed and disposed by the reader
            { 
                var stream = Configuration.GetTestWorkbook("xTest10x10");
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    var result = excelReader.AsDataSet();
                }

                Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
            }

            // Verify streams used by standard encryption are closed
            {
                var stream = Configuration.GetTestWorkbook("standard_AES128_SHA1_ECB_pwd_password");

                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(
                    stream,
                    new ExcelReaderConfiguration() { Password = "password" }))
                {
                    var result = excelReader.AsDataSet();
                }

                Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
            }

            // Verify streams used by agile encryption are closed
            {
                var stream = Configuration.GetTestWorkbook("agile_AES128_MD5_CBC_pwd_password");

                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(
                    stream,
                    new ExcelReaderConfiguration() { Password = "password" }))
                {
                    var result = excelReader.AsDataSet();
                }

                Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
            }
        }

        [TestMethod]
        public void GitIssue_271_InvalidDimension()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_271_InvalidDimension")))
            {
                var dataset = excelReader.AsDataSet();
                Assert.AreEqual(3, dataset.Tables[0].Columns.Count);
                Assert.AreEqual(9, dataset.Tables[0].Rows.Count);
            }
        }

        [TestMethod]
        public void GitIssue_283_TimeSpan()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest_git_issue_283_TimeSpan")))
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
        public void GitIssue_289_CompoundDocumentEncryptedWithDefaultPassword()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue289")))
            {
                reader.Read();
                Assert.AreEqual("aaaaaaa", reader.GetValue(0));
            }
        }

        [TestMethod]
        public void MergedCells()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_MergedCell_OpenXml")))
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
                    new int[]
                    {
                        1,
                        2,
                        0,
                        1
                    },
                    new int[]
                    {
                        mergedCells[0].FromRow,
                        mergedCells[0].ToRow,
                        mergedCells[0].FromColumn,
                        mergedCells[0].ToColumn
                    }
                );

                CollectionAssert.AreEqual(
                    new int[]
                    {
                        1,
                        5,
                        2,
                        2
                    },
                    new int[]
                    {
                        mergedCells[1].FromRow,
                        mergedCells[1].ToRow,
                        mergedCells[1].FromColumn,
                        mergedCells[1].ToColumn
                    }
                );

                CollectionAssert.AreEqual(
                    new int[]
                    {
                        3,
                        5,
                        0,
                        0
                    },
                    new int[]
                    {
                        mergedCells[2].FromRow,
                        mergedCells[2].ToRow,
                        mergedCells[2].FromColumn,
                        mergedCells[2].ToColumn
                    }
                );

                CollectionAssert.AreEqual(
                    new int[]
                    {
                        6,
                        6,
                        0,
                        2
                    },
                    new int[]
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
        public void GitIssue_301_IgnoreCase()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_301_IgnoreCase")))
            {
                DataTable result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(10, result.Rows.Count);
                Assert.AreEqual(10, result.Columns.Count);
                Assert.AreEqual("10x10", result.Rows[1][0]);
                Assert.AreEqual("10x27", result.Rows[9][9]);
            }
        }

        [TestMethod]
        public void GitIssue_319_InlineRichText()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue319")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual("Text1", result.Rows[0][0]);
            }
        }

        [TestMethod]
        public void GitIssue_324_MultipleRowElementsPerRow()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_324")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(20, result.Rows.Count);
                Assert.AreEqual(13, result.Columns.Count);

                Assert.That(result.Rows[10].ItemArray, Is.EqualTo(new object[] { DBNull.Value, DBNull.Value, "Other", 191036.15, 194489.45, 66106.32, 37167.88, 102589.54, 57467.94, 130721.93, 150752.67, 76300.69, 67024.6 }));
            }
        }

        [TestMethod]
        public void GitIssue_323_DoubleClose()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("xTest10x10")))
            {
                reader.Read();
                reader.Close();
            }
        }

        [TestMethod]
        public void GitIssue_329_Error()
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
    }
}