using System;
#if NET20 || NET45 || NETCOREAPP2_0
using System.Data;
#endif
using System.IO;
using ExcelDataReader.Exceptions;
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
        public void GitIssue_70_ExcelBinaryReader_tryConvertOADateTime_forumla()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_70")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);

                var date = ds.Tables[0].Rows[1].ItemArray[0];

                Assert.AreEqual(new DateTime(2014, 01, 01), date);
            }
        }

        [TestMethod]
        public void GitIssue_51_ReadCellLabel()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_51")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);

                var value = ds.Tables[0].Rows[0].ItemArray[1];

                Assert.AreEqual("Monetary aggregates (R millions)", value);
            }
        }

        [TestMethod]
        public void GitIssue_29_ReadSheetStatesReadsCorrectly()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Excel_Dataset")))
            {
                Assert.AreEqual("hidden", excelReader.VisibleState);

                excelReader.NextResult();
                Assert.AreEqual("visible", excelReader.VisibleState);

                excelReader.NextResult();
                Assert.AreEqual("veryhidden", excelReader.VisibleState);
            }
        }

        [TestMethod]
        public void GitIssue_29_AsDataSetProvidesCorrectSheetVisibleState()
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Excel_Dataset")))
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
        public void GitIssue_45()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_45")))
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
        public void AsDataSet_Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestChess")))
            {
                DataSet result = excelReader.AsDataSet();

                Assert.IsTrue(result != null);
                Assert.AreEqual(1, result.Tables.Count);
                Assert.AreEqual(4, result.Tables[0].Rows.Count);
                Assert.AreEqual(6, result.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void AsDataSet_Test_Row_Count()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestChess")))
            {
                var result = excelReader.AsDataSet(Configuration.NoColumnNamesConfiguration);

                Assert.AreEqual(4, result.Tables[0].Rows.Count);
            }
        }

        [TestMethod]
        public void AsDataSet_Test_Row_Count_FirstRowAsColumnNames()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestChess")))
            {
                var result = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(3, result.Tables[0].Rows.Count);
            }
        }

        [TestMethod]
        public void Issue_11553_11570_FATIssue_Offset()
        {
            void DoTestFATStreamIssue(string sheetId)
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

            void DoTestFATStreamIssueType2(string sheetId)
            {
                var filePath = Configuration.GetTestWorkbookPath(sheetId);

                using (Stream stream = new MemoryStream(File.ReadAllBytes(filePath)))
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream))
                {
                    // ReSharper disable once AccessToDisposedClosure
                    Assert.DoesNotThrow(() => excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration));
                }
            }

            DoTestFATStreamIssue("Test_Issue_11553_FAT");
            DoTestFATStreamIssueType2("Test_Issue_11570_FAT_1");
            DoTestFATStreamIssueType2("Test_Issue_11570_FAT_2");
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestChess")))
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
            using (IExcelDataReader r = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestMultiSheet")))
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
                Assert.AreEqual(2, fieldCount);
                Assert.AreEqual(3, table.Rows[4][1]);

                Assert.AreEqual(false, r.NextResult());
            }
        }

        [TestMethod]
        public void DataReader_Read_Test()
        {
            using (IExcelDataReader r = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_num_double_date_bool_string")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test10x10000")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test10x10")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test255x10")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestDoublePrecision")))
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
        public void Fail_Test()
        {
            var exception = Assert.Throws<HeaderException>(() =>
            {
                using (ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestFail_Binary")))
                {
                }
            });

            Assert.AreEqual("Invalid file signature.", exception.Message);
        }

        [TestMethod]
        public void Issue_Date_and_Time_1468_Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Encoding_1520")))
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
        public void Issue_8536_Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_8536")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11397")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_4031_NullColumn")))
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
        public void Issue_10725()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_10725")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11435_Colors")))
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
        public void Issue_7433_IllegalOleAutDate()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_7433_IllegalOleAutDate")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_BoolFormula")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                Assert.AreEqual(true, dataSet.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void Issue_Decimal_1109_Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Decimal_1109")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Encoding_1520")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestMultiSheet")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_num_double_date_bool_string")))
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
        public void Issue_11479_BlankSheet()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Issue_11479_BlankSheet")))
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
        public void Test_BlankHeader()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_BlankHeader")))
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
        public void Test_OpenOffice()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_OpenOffice")))
            {
                AssertUtilities.DoOpenOfficeTest(excelReader);
            }
        }

        /// <summary>
        /// Issue 11 - OpenOffice files were skipping the first row if IsFirstRowAsColumnNames = false;
        /// </summary>
        [TestMethod]
        public void GitIssue_11_OpenOffice_Row_Count()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_OpenOffice")))
            {
                var dataset = excelReader.AsDataSet(Configuration.NoColumnNamesConfiguration);
                Assert.AreEqual(34, dataset.Tables[0].Rows.Count);
            }
        }

        /// <summary>
        /// This test is to ensure that we get the same results from an xls saved in excel vs open office
        /// </summary>
        [TestMethod]
        public void Test_OpenOffice_SavedInExcel()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Excel_OpenOffice")))
            {
                AssertUtilities.DoOpenOfficeTest(excelReader);
            }
        }

        [TestMethod]
        public void Test_Issue_11601_ReadSheetnames()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Excel_Dataset")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestUnicodeChars")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Uncalculated")))
            {
                var dataset = excelReader.AsDataSet();
                Assert.IsNotNull(dataset);
                Assert.AreNotEqual(dataset.Tables.Count, 0);
                var table = dataset.Tables[0];
                Assert.IsNotNull(table);

                Assert.AreEqual("1", table.Rows[1][0].ToString());
                Assert.AreEqual("3", table.Rows[1][2].ToString());
                Assert.AreEqual("3", table.Rows[1][4].ToString());
            }
        }

        [TestMethod]
        public void Issue_11570_Excel2013()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11570_Excel2013")))
            {
                var dataset = excelReader.AsDataSet();

                Assert.AreEqual(2, dataset.Tables[0].Columns.Count);
                Assert.AreEqual(5, dataset.Tables[0].Rows.Count);

                Assert.AreEqual("1.1.1.2", dataset.Tables[0].Rows[0][0]);
                Assert.AreEqual(10d, dataset.Tables[0].Rows[0][1]);

                Assert.AreEqual("1.1.1.15", dataset.Tables[0].Rows[1][0]);
                Assert.AreEqual(3d, dataset.Tables[0].Rows[1][1]);

                Assert.AreEqual("2.1.2.23", dataset.Tables[0].Rows[2][0]);
                Assert.AreEqual(14d, dataset.Tables[0].Rows[2][1]);

                Assert.AreEqual("2.1.2.31", dataset.Tables[0].Rows[3][0]);
                Assert.AreEqual(2d, dataset.Tables[0].Rows[3][1]);

                Assert.AreEqual("2.8.7.30", dataset.Tables[0].Rows[4][0]);
                Assert.AreEqual(2d, dataset.Tables[0].Rows[4][1]);
            }
        }

        [TestMethod]
        public void Issue_11572_CodePage()
        {
            // This test was skipped for a long time as it produced: "System.NotSupportedException : No data is available for encoding 27651."
            // Upon revisting the underlying cause appears to be fixed
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11572_CodePage")))
            {
                Assert.DoesNotThrow(() => excelReader.AsDataSet());
            }
        }

        /// <summary>
        /// Not fixed yet
        /// </summary>
        [TestMethod]
        public void Issue_11545_NoIndex()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11545_NoIndex")))
            {
                var dataset = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual("CI2229         ", dataset.Tables[0].Rows[0][0]);
                Assert.AreEqual("12069E01018A1  ", dataset.Tables[0].Rows[0][6]);
                Assert.AreEqual(new DateTime(2012, 03, 01), dataset.Tables[0].Rows[0][8]);
            }
        }

        [TestMethod]
        public void Issue_11573_BlankValues()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11573_BlankValues")))
            {
                var dataset = excelReader.AsDataSet();

                Assert.AreEqual(1D, dataset.Tables[0].Rows[12][0]);
                Assert.AreEqual("070202", dataset.Tables[0].Rows[12][1]);
            }
        }

        [TestMethod]
        public void Issue_DateFormatButNotDate()
        {
            // we want to make sure that if a cell is formatted as a date but it's contents are not a date then
            // the output is not a date
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_DateFormatButNotDate")))
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
        public void Issue_11642_ValuesNotLoaded()
        {
            // Excel.Log.Log.InitializeWith<Log4NetLog>();
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11642_ValuesNotLoaded")))
            {
                var dataset = excelReader.AsDataSet();

                Assert.AreEqual("431113*", dataset.Tables[2].Rows[29][1].ToString());
                Assert.AreEqual("024807", dataset.Tables[2].Rows[36][1].ToString());
                Assert.AreEqual("160019", dataset.Tables[2].Rows[53][1].ToString());
            }
        }

        [TestMethod]
        public void Issue_11636_BiffStream()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11636_BiffStream")))
            {
                var dataset = excelReader.AsDataSet();

                // check a couple of values
                Assert.AreEqual("SP011", dataset.Tables[0].Rows[9][0]);
                Assert.AreEqual(9.9, dataset.Tables[0].Rows[32][11]);
                Assert.AreEqual(78624.44, dataset.Tables[1].Rows[27][12]);
            }
        }

        /// <summary>
        /// Not fixed yet
        /// The problem occurs with unseekable stream and loigc related to minifat that uses seek
        /// It should probably only use seek if it needs to go backwards, I think at the moment it uses seek all the time
        /// which is probably not good for performance
        /// </summary>
        [TestMethod]
        [Ignore("Not fixed yet")]
        public void Issue_11639_11644_ForwardOnlyStream()
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
        /// The problem occurs with unseekable stream and loigc related to minifat that uses seek
        /// It should probably only use seek if it needs to go backwards, I think at the moment it uses seek all the time
        /// which is probably not good for performance
        /// </summary>
        [TestMethod]
        public void Issue_12556_corrupt()
        {
            Assert.Throws<InvalidOperationException>(() =>
            {
                // Excel.Log.Log.InitializeWith<Log4NetLog>();
                using (var forwardStream = Configuration.GetTestWorkbook("Test_Issue_12556_corrupt"))
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
        public void Test_Issue_11818_OutOfRange()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11818_OutOfRange")))
            {
                var dataset = excelReader.AsDataSet();

                Assert.AreEqual("Total Revenue", dataset.Tables[0].Rows[10][0]);
            }
        }

        [TestMethod]
        public void Test_Issue_111_NoRowRecords()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_111_NoRowRecords")))
            {
                var dataset = excelReader.AsDataSet();

                Assert.AreEqual(1, dataset.Tables.Count);
                Assert.AreEqual(12, dataset.Tables[0].Rows.Count);
                Assert.AreEqual(14, dataset.Tables[0].Columns.Count);

                Assert.AreEqual(2015.0, dataset.Tables[0].Rows[7][0]);
            }
        }

        [TestMethod]
        public void Test_Git_Issue_145()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_145")))
            {
                excelReader.Read();
                excelReader.Read();
                excelReader.Read();

                string value = excelReader.GetString(3);

                Assert.AreEqual("Japanese Government Bonds held by the Bank of Japan", value);
            }
        }

        [TestMethod]
        public void Test_GitIssue_152_SheetName_UTF16LE_Compressed()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_152")))
            {
                var dataset = excelReader.AsDataSet();

                Assert.AreEqual("åäöñ", dataset.Tables[0].TableName);
            }
        }

        [TestMethod]
        public void Test_GitIssue_152_Cell_UTF16LE_Compressed()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_152")))
            {
                var dataset = excelReader.AsDataSet();

                Assert.AreEqual("åäöñ", dataset.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void GitIssue_158()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_158")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);

                var date = ds.Tables[0].Rows[3].ItemArray[2];

                Assert.AreEqual(new DateTime(2016, 09, 10), date);
            }
        }

        [TestMethod]
        public void GitIssue_173()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_173")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);
                Assert.AreEqual(40, ds.Tables.Count);
            }
        }

        [TestMethod]
        public void ReadWriteProtectedStructureUsingStandardEncryption()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("protectedsheet-xxx")))
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
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestTableOnlyImage_x01oct2016")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);
                Assert.AreEqual(4, ds.Tables.Count);
            }
        }

        [TestMethod]
        public void AllowFfffAsByteOrder()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_InvalidByteOrderValueInHeader")))
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
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("AllColumnsNotReadInHiddenTable")))
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
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("RowWithDifferentNumberOfColumns")))
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
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_145")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual(256, ds.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void Row1217NotRead()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Row1217NotRead")))
            {
                var ds = excelReader.AsDataSet();
                CollectionAssert.AreEqual(new object[] { DBNull.Value, "Año", "Mes", DBNull.Value, "Índice", "Variación Mensual", "Variación Acumulada", "Variación en 12 Meses", "Incidencia Mensual", "Incidencia Acumulada", "Incidencia a 12 Meses", DBNull.Value, DBNull.Value }, ds.Tables[0].Rows[1216].ItemArray);
            }
        }

        [TestMethod]
        public void StringContinuationAfterCharacterData()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("StringContinuationAfterCharacterData")))
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
            using (var stream = Configuration.GetTestWorkbook("biff3"))
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
            using (var stream = Configuration.GetTestWorkbook("Test_git_issue_5"))
                Assert.Throws<InvalidOperationException>(() => ExcelReaderFactory.CreateBinaryReader(stream));
        }

        [TestCase]
        public void Issue2InvalidDimensionRecord()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_2")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual(new[] { "A1", "B1" }, ds.Tables[0].Rows[0].ItemArray);
            }
        }

        [TestCase]
        public void ExcelLibrary_NonContinousMiniStream()
        {
            // Verify the output from the sample code for the ExcelLibrary package parses
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("ExcelLibrary_newdoc")))
            {
                Assert.DoesNotThrow(() => excelReader.AsDataSet());
            }
        }

        [TestCase]
        public void GitIssue_184_AdditionalFATSectors()
        {
            // Big spreadsheets have additional sectors beyond the header with FAT contents
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("GitIssue_184_FATSectors")))
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
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_217")))
            {
                var ds = excelReader.AsDataSet();
                CollectionAssert.AreEqual(new object[] { "REX GESAMT      ", 484.7929, 142.1032, -0.1656, 5.0315225293000001, 5.0398685515999997, 37.5344725251, DBNull.Value, DBNull.Value }, ds.Tables[2].Rows[10].ItemArray);
            }
        }

        [TestMethod]
        public void GitIssue_231_NoCodePage()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_231_NoCodePage")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual(11, ds.Tables[0].Columns.Count);
                Assert.AreEqual(5, ds.Tables[0].Rows.Count);
            }
        }

        [TestMethod]
        public void GitIssue_82_Date1900_Binary()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("roo_1900_base")))
            {
                // 15/06/2009
                // 28/06/2009 (=TODAY() when file was saved)

                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(new DateTime(2009, 6, 15), (DateTime)result.Tables[0].Rows[0][0]);
                Assert.AreEqual(new DateTime(2009, 6, 28), (DateTime)result.Tables[0].Rows[1][0]);
            }
        }

        [TestMethod]
        public void GitIssue_82_Date1904_Binary()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("roo_1904_base")))
            {
                // 15/06/2009
                // 28/06/2009 (=TODAY() when file was saved)

                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(new DateTime(2009, 6, 15), (DateTime)result.Tables[0].Rows[0][0]);
                Assert.AreEqual(new DateTime(2009, 6, 28), (DateTime)result.Tables[0].Rows[1][0]);
            }
        }

        [TestMethod]
        public void As3xls_BIFF2()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF2")))
            {
                DataSet result = excelReader.AsDataSet();
                Test_As3xls(result);
            }
        }

        [TestMethod]
        public void As3xls_BIFF3()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF3")))
            {
                DataSet result = excelReader.AsDataSet();
                Test_As3xls(result);
            }
        }

        [TestMethod]
        public void As3xls_BIFF4()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF4")))
            {
                DataSet result = excelReader.AsDataSet();
                Test_As3xls(result);
            }
        }

        [TestMethod]
        public void As3xls_BIFF5()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF5")))
            {
                DataSet result = excelReader.AsDataSet();
                Test_As3xls(result);
            }
        }

        void Test_As3xls(DataSet result)
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
        public void GitIssue_240_ExceptionBeforeRead()
        {
            // Check the exception and message when trying to get data before calling Read().
            // Using the same as SqlDataReader, making it easier to search for a general solution.
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test10x10")))
            {
                var exception = Assert.Throws<InvalidOperationException>(() =>
                {
                    for (int columnIndex = 0; columnIndex < excelReader.FieldCount; columnIndex++)
                    {
                        string name = excelReader.GetString(columnIndex);
                    }
                });

                Assert.AreEqual("No data exists for the row/column.", exception.Message);
            }
        }

        [TestMethod]
        public void GitIssue_241_Simple()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_224_simple_biff")))
            {
                Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft åäö &T&CCenter åäö &D&RRight  åäö &P"), "Header");
                Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft åäö &P&CFooter åäö &P&RRight åäö &D"), "Footer");
            }
        }

        [TestMethod]
        public void GitIssue_241_Simple95()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_224_simple_biff95")))
            {
                Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft åäö &T&CCenter åäö &D&RRight  åäö &P"), "Header");
                Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft åäö &P&CFooter åäö &P&RRight åäö &D"), "Footer");
            }
        }

        [TestMethod]
        public void GitIssue_245_CodeName()
        {
            // Test no CodeName = null
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test10x10")))
            {
                Assert.AreEqual(null, reader.CodeName);
            }

            // Test CodeName is set
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Excel_Dataset")))
            {
                Assert.AreEqual("Sheet1", reader.CodeName);
            }

            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_45")))
            {
                Assert.AreEqual("Hoja8", reader.CodeName);
            }
        }

        [TestMethod]
        public void GitIssue_250_RichText()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_250_richtext")))
            {
                reader.Read();
                var text = reader.GetString(0);
                Assert.AreEqual("Lorem ipsum dolor sit amet, ei pri verterem efficiantur, per id meis idque deterruisset.", text);
            }
        }

        [TestMethod]
        public void GitIssue_242_Password()
        {
            // BIFF8 standard encryption cryptoapi rc4+sha 
            using (var reader = ExcelReaderFactory.CreateBinaryReader(
                Configuration.GetTestWorkbook("Test_git_issue_242_std_rc4_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // Pre-BIFF8 xor obfuscation
            using (var reader = ExcelReaderFactory.CreateBinaryReader(
                Configuration.GetTestWorkbook("Test_git_issue_242_xor_pwd_password"),
                new ExcelReaderConfiguration() { Password = "password" }))
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
                    Configuration.GetTestWorkbook("Test_git_issue_242_xor_pwd_password"),
                    new ExcelReaderConfiguration() { Password = "wrongpassword" }))
                {
                    reader.Read();
                }
            });

            Assert.Throws<InvalidPasswordException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateBinaryReader(
                    Configuration.GetTestWorkbook("Test_git_issue_242_xor_pwd_password")))
                {
                    reader.Read();
                }
            });
        }

        [TestMethod]
        public void GitIssue_263()
        {
            using (var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test_git_issue_263")))
            {
                var ds = reader.AsDataSet();
                Assert.AreEqual("Economic Inactivity by age\n(Official statistics: not designated as National Statistics)", ds.Tables[1].Rows[3][0]);
            }
        }

        [TestMethod]
        public void BinaryRowHeight()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("CollapsedHide")))
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
        public void GitIssue_270_EmptyRowsAtTheEnd()
        {
            // AsDataSet() trims trailing blank rows
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_270")))
            {
                var dataset = reader.AsDataSet();
                Assert.AreEqual(1, dataset.Tables[0].Rows.Count);
            }

            // Reader methods do not trim trailing blank rows
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_270")))
            {
                var rowCount = 0;
                while (reader.Read())
                    rowCount++;
                Assert.AreEqual(65536, rowCount);
            }
        }

        static bool IsEmptyOrHiddenRow(IExcelDataReader reader)
        {
            if (reader.RowHeight == 0)
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
        public void GitIssue_160_FilterRow()
        {
            // Check there are four rows with data, including empty and hidden rows
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("CollapsedHide")))
            {
                var dataset = reader.AsDataSet();

                Assert.AreEqual(4, dataset.Tables[0].Rows.Count);
            }

            // Check there are two rows with content
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("CollapsedHide")))
            {
                var dataset = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                    {
                        FilterRow = rowReader => !IsEmptyRow(rowReader)
                    }
                });

                Assert.AreEqual(2, dataset.Tables[0].Rows.Count);
            }

            // Check there is one visible row with content
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("CollapsedHide")))
            {
                var dataset = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                    {
                        FilterRow = rowReader => !IsEmptyOrHiddenRow(rowReader)
                    }
                });

                Assert.AreEqual(1, dataset.Tables[0].Rows.Count);
            }
        }

        [TestMethod]
        public void GitIssue_265_BinaryDisposed()
        {
            var stream = Configuration.GetTestWorkbook("Test10x10");
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream))
            {
                var result = excelReader.AsDataSet();
            }

            Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
        }

        [TestMethod]
        public void GitIssue_286_SSTStringHeader()
        {
            // Parse xls with SST containing string split exactly between its header and string data across the BIFF Continue records
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_286_SST")))
            {
                Assert.IsNotNull(reader);
            }
        }
    }
}