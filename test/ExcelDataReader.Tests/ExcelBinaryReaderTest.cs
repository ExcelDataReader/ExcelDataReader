using System;
#if !NETCOREAPP1_0
using System.Data;
#endif
using System.Runtime.InteropServices.ComTypes;
using Excel;
using System.IO;

using NUnit.Framework;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestInitialize = NUnit.Framework.SetUpAttribute;
using TestCleanup = NUnit.Framework.TearDownAttribute;
using TestMethod = NUnit.Framework.TestAttribute;

namespace ExcelDataReader.Tests
{

	[TestClass]
    
    public class ExcelBinaryReaderTest
    {
		[TestInitialize]
		public void TestInitialize()
		{
#if NETCOREAPP1_0
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif
		}

		[TestMethod]
        public void GitIssue_70_ExcelBinaryReader_tryConvertOADateTime_forumla()
        {
            var excelReader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Git_Issue_70"), true);

            var ds = excelReader.AsDataSet(true);
            Assert.IsNotNull(ds);

            var date = ds.Tables[0].Rows[1].ItemArray[0];

            Assert.AreEqual(new DateTime(2014,01,01), date);
        }

        [TestMethod]
        public void GitIssue_51_ReadCellLabel()
        {
            var excelReader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Git_Issue_51"), true);

            var ds = excelReader.AsDataSet(true);
            Assert.IsNotNull(ds);

            var value = ds.Tables[0].Rows[0].ItemArray[1];

            Assert.AreEqual("Monetary aggregates (R millions)", value);
        }

        [TestMethod]
        public void GitIssue_29_ReadSheetStatesReadsCorrectly()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Excel_Dataset"));

            Assert.AreEqual("hidden", excelReader.VisibleState);

            excelReader.NextResult();
            Assert.AreEqual("visible", excelReader.VisibleState);

            excelReader.NextResult();
            Assert.AreEqual("veryhidden", excelReader.VisibleState);
        }

        [TestMethod]
        public void GitIssue_29_AsDataSetProvidesCorrectSheetVisibleState()
        {
            IExcelDataReader reader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Excel_Dataset"));

            var dataset = reader.AsDataSet();

            reader.Close();

            Assert.IsTrue(dataset != null);
            Assert.AreEqual(3, dataset.Tables.Count);
            Assert.AreEqual("hidden", dataset.Tables[0].ExtendedProperties["visiblestate"]);
            Assert.AreEqual("visible", dataset.Tables[1].ExtendedProperties["visiblestate"]);
            Assert.AreEqual("veryhidden", dataset.Tables[2].ExtendedProperties["visiblestate"]);
        }

        [TestMethod]
        public void GitIssue_45()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_git_issue_45")))
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
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("TestChess"));

            DataSet result = excelReader.AsDataSet();

            Assert.IsTrue(result != null);
            Assert.AreEqual(1, result.Tables.Count);
            Assert.AreEqual(4, result.Tables[0].Rows.Count);
            Assert.AreEqual(6, result.Tables[0].Columns.Count);

            excelReader.Close();
        }

        [TestMethod]
        public void AsDataSet_Test_Row_Count()
        {
            var excelReader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("TestChess"));
            excelReader.IsFirstRowAsColumnNames = false;
            var result = excelReader.AsDataSet();

            Assert.AreEqual(4, result.Tables[0].Rows.Count);


            excelReader.Close();
        }

        [TestMethod]
        public void AsDataSet_Test_Row_Count_FirstRowAsColumnNames()
        {
            var excelReader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("TestChess"));
            excelReader.IsFirstRowAsColumnNames = true;
            var result = excelReader.AsDataSet();

            Assert.AreEqual(3, result.Tables[0].Rows.Count);


            excelReader.Close();
        }

        [TestMethod]
        public void Issue_11553_11570_FATIssue_Offset()
        {

			DoTestFATStreamIssue("Test_Issue_11553_FAT");
			DoTestFATStreamIssueType2("Test_Issue_11570_FAT_1");
			DoTestFATStreamIssueType2("Test_Issue_11570_FAT_2");

        }

	    private static void DoTestFATStreamIssue(string sheetId)
	    {
		    var excelReader1 = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook(sheetId)); // Works.

		    string filePath = Helper.GetTestWorkbookPath(sheetId);
		    Assert.IsNotNull(excelReader1);

		    var ms1 = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read,
		                                       System.IO.FileShare.ReadWrite);
		    var excelReader2 = ExcelReaderFactory.CreateBinaryReader(ms1); // Works!
		    Assert.IsNotNull(excelReader2);

		    var bytes = System.IO.File.ReadAllBytes(filePath);
		    var ms2 = new System.IO.MemoryStream(bytes);

		    var excelReader3 = ExcelReaderFactory.CreateBinaryReader(ms2); // Did not work, but does now
		    Assert.IsNotNull(excelReader3);
	    }

		private static void DoTestFATStreamIssueType2(string sheetId)
		{
			var filePath = Helper.GetTestWorkbookPath(sheetId);
			
			using (Stream stream = new MemoryStream(File.ReadAllBytes(filePath)))
			{
				IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

				//Options
				excelReader.IsFirstRowAsColumnNames = true;
				DataSet result = excelReader.AsDataSet();
				excelReader.Close();

			}
		}

	    //[TestMethod]
        //public void Test_SSRS()
        //{
        //	IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_SSRS"));

        //	DataSet result = excelReader.AsDataSet();



        //	excelReader.Close();
        //}


        [TestMethod]
        public void ChessTest()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("TestChess"));

            DataTable result = excelReader.AsDataSet().Tables[0];

            Assert.AreEqual(4, result.Rows.Count);
            Assert.AreEqual(6, result.Columns.Count);
            Assert.AreEqual("1", result.Rows[3][5].ToString());
            Assert.AreEqual("1", result.Rows[2][0].ToString());

            excelReader.Close();
        }

        [TestMethod]
        public void DataReader_NextResult_Test()
        {
            IExcelDataReader r = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("TestMultiSheet"));

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
                table.Rows.Add(r.GetInt32(0), r.GetInt32(1), r.GetInt32(2), r.GetInt32(3));
            }

            Assert.AreEqual(12, table.Rows.Count);
            Assert.AreEqual(4, fieldCount);
            Assert.AreEqual(1, table.Rows[11][3]);


            r.NextResult();
            table.Rows.Clear();

            while (r.Read())
            {
                fieldCount = r.FieldCount;
                table.Rows.Add(r.GetInt32(0), r.GetInt32(1), r.GetInt32(2), r.GetInt32(3));
            }

            Assert.AreEqual(12, table.Rows.Count);
            Assert.AreEqual(4, fieldCount);
            Assert.AreEqual(2, table.Rows[11][3]);


            r.NextResult();
            table.Rows.Clear();

            while (r.Read())
            {
                fieldCount = r.FieldCount;
                table.Rows.Add(r.GetInt32(0), r.GetInt32(1));
            }

            Assert.AreEqual(5, table.Rows.Count);
            Assert.AreEqual(2, fieldCount);
            Assert.AreEqual(3, table.Rows[4][1]);

            Assert.AreEqual(false, r.NextResult());

            r.Close();
        }

        [TestMethod]
        public void DataReader_Read_Test()
        {
            IExcelDataReader r =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_num_double_date_bool_string"));

            var table = new DataTable();
            table.Columns.Add(new DataColumn("num_col", typeof(int)));
            table.Columns.Add(new DataColumn("double_col", typeof(double)));
            table.Columns.Add(new DataColumn("date_col", typeof(DateTime)));
            table.Columns.Add(new DataColumn("boo_col", typeof(bool)));

            int fieldCount = -1;

            while (r.Read())
            {
                fieldCount = r.FieldCount;
                table.Rows.Add(r.GetInt32(0), r.GetDouble(1), r.GetDateTime(2), r.IsDBNull(4));
            }

            r.Close();

            Assert.AreEqual(6, fieldCount);

            Assert.AreEqual(30, table.Rows.Count);

            Assert.AreEqual(1, int.Parse(table.Rows[0][0].ToString()));
            Assert.AreEqual(1346269, int.Parse(table.Rows[29][0].ToString()));

            //double + Formula
            Assert.AreEqual(1.02, double.Parse(table.Rows[0][1].ToString()));
            Assert.AreEqual(4.08, double.Parse(table.Rows[2][1].ToString()));
            Assert.AreEqual(547608330.24, double.Parse(table.Rows[29][1].ToString()));

            //Date + Formula
            Assert.AreEqual(new DateTime(2009, 5, 11).ToShortDateString(),
                            DateTime.Parse(table.Rows[0][2].ToString()).ToShortDateString());
            Assert.AreEqual(new DateTime(2009, 11, 30).ToShortDateString(),
                            DateTime.Parse(table.Rows[29][2].ToString()).ToShortDateString());
        }

        [TestMethod]
        public void Dimension10x10000Test()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test10x10000"));

            DataTable result = excelReader.AsDataSet().Tables[0];

            Assert.AreEqual(10000, result.Rows.Count);
            Assert.AreEqual(10, result.Columns.Count);
            Assert.AreEqual("1x2", result.Rows[1][1]);
            Assert.AreEqual("1x10", result.Rows[1][9]);
            Assert.AreEqual("1x1", result.Rows[9999][0]);
            Assert.AreEqual("1x10", result.Rows[9999][9]);

            excelReader.Close();
        }

        [TestMethod]
        public void Dimension10x10Test()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test10x10"));

            DataTable result = excelReader.AsDataSet().Tables[0];

            Assert.AreEqual(10, result.Rows.Count);
            Assert.AreEqual(10, result.Columns.Count);
            Assert.AreEqual("10x10", result.Rows[1][0]);
            Assert.AreEqual("10x27", result.Rows[9][9]);

            excelReader.Close();
        }

        [TestMethod]
        public void Dimension255x10Test()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test255x10"));

            DataTable result = excelReader.AsDataSet().Tables[0];

            Assert.AreEqual(10, result.Rows.Count);
            Assert.AreEqual(255, result.Columns.Count);
            Assert.AreEqual("1", result.Rows[9][254].ToString());
            Assert.AreEqual("one", result.Rows[1][1].ToString());

            excelReader.Close();
        }

        [TestMethod]
        public void DoublePrecisionTest()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("TestDoublePrecision"));

            DataTable result = excelReader.AsDataSet().Tables[0];

            const double excelPI = 3.14159265358979;

            Assert.AreEqual(+excelPI, result.Rows[2][1]);
            Assert.AreEqual(-excelPI, result.Rows[3][1]);

            Assert.AreEqual(+excelPI * 1.0e-300, result.Rows[4][1]);
            Assert.AreEqual(-excelPI * 1.0e-300, result.Rows[5][1]);

            Assert.AreEqual(+excelPI * 1.0e300, (double)result.Rows[6][1], 1e286); //only accurate to 1e286 because excel only has 15 digits precision
			Assert.AreEqual(-excelPI * 1.0e300, (double)result.Rows[7][1], 1e286);

            Assert.AreEqual(+excelPI * 1.0e14, result.Rows[8][1]);
            Assert.AreEqual(-excelPI * 1.0e14, result.Rows[9][1]);

            excelReader.Close();
        }

        [TestMethod]
        public void Fail_Test()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("TestFail_Binary"));

            Assert.AreEqual(false, excelReader.IsValid);
            Assert.AreEqual(true, excelReader.IsClosed);
            Assert.AreEqual("Error: Invalid file signature.", excelReader.ExceptionMessage);
        }


        [TestMethod]
        public void Issue_Date_and_Time_1468_Test()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Encoding_1520"));

            DataSet dataSet = excelReader.AsDataSet(true);

            string val1 = new DateTime(2009, 05, 01).ToShortDateString();
            string val2 = DateTime.Parse(dataSet.Tables[0].Rows[1][1].ToString()).ToShortDateString();

            Assert.AreEqual(val1, val2);

            val1 = DateTime.Parse("11:00:00").ToShortTimeString();
            val2 = DateTime.Parse(dataSet.Tables[0].Rows[2][4].ToString()).ToShortTimeString();

            Assert.AreEqual(val1, val2);

            excelReader.Close();
        }

        [TestMethod]
        public void Issue_8536_Test()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_8536"));

            DataSet dataSet = excelReader.AsDataSet(true);


            //date
            var date1900 = dataSet.Tables[0].Rows[7][1];
            Assert.AreEqual(typeof(DateTime), date1900.GetType());
            Assert.AreEqual(new DateTime(1900, 1, 1), date1900);

            //xml encoded chars
            var xmlChar1 = dataSet.Tables[0].Rows[6][1];
            Assert.AreEqual(typeof(string), xmlChar1.GetType());
            Assert.AreEqual("&#x26; ", xmlChar1);

            //number but matches a date serial
            var datenum1 = dataSet.Tables[0].Rows[5][1];
            Assert.AreEqual(typeof(double), datenum1.GetType());
            Assert.AreEqual(41244, double.Parse(datenum1.ToString()));

            //date
            var date1 = dataSet.Tables[0].Rows[4][1];
            Assert.AreEqual(typeof(DateTime), date1.GetType());
            Assert.AreEqual(new DateTime(2012, 12, 1), date1);

            //double
            var num1 = dataSet.Tables[0].Rows[3][1];
            Assert.AreEqual(typeof(double), num1.GetType());
            Assert.AreEqual(12345, double.Parse(num1.ToString()));

            //boolean issue
            var val2 = dataSet.Tables[0].Rows[2][1];
            Assert.AreEqual(typeof(bool), val2.GetType());
            Assert.IsTrue((bool)val2);




            excelReader.Close();
        }

        [TestMethod]
        public void Issue_11397_Currency_Test()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_11397"));

            DataSet dataSet = excelReader.AsDataSet(true);

            excelReader.Close();

            Assert.AreEqual("$44.99", dataSet.Tables[0].Rows[1][0].ToString()); //general in spreadsheet so should be a string including the $
            Assert.AreEqual(44.99, double.Parse(dataSet.Tables[0].Rows[2][0].ToString())); //currency euros in spreadsheet so should be a currency
            Assert.AreEqual(44.99, double.Parse(dataSet.Tables[0].Rows[3][0].ToString())); //currency pounds in spreadsheet so should be a currency

        }

        [TestMethod]
        public void Issue_4031_NullColumn()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_4031_NullColumn"));

            //DataSet dataSet = excelReader.AsDataSet(true);

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
            Assert.AreEqual(1, excelReader.GetInt32(4));

            excelReader.Close();


        }

        [TestMethod]
        public void Issue_10725()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_10725"));


            excelReader.Read();
            Assert.AreEqual(8.8, excelReader.GetValue(0));

            DataSet result = excelReader.AsDataSet();

            Assert.AreEqual(8.8, result.Tables[0].Rows[0][0]);

            excelReader.Close();


        }

        [TestMethod]
        public void Issue_11435_Colors()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_11435_Colors"));

            DataSet dataSet = excelReader.AsDataSet(true);

            Assert.AreEqual("test1", dataSet.Tables[0].Rows[0][0].ToString());
            Assert.AreEqual("test2", dataSet.Tables[0].Rows[0][1].ToString());
            Assert.AreEqual("test3", dataSet.Tables[0].Rows[0][2].ToString());


            excelReader.Read();

            Assert.AreEqual("test1", excelReader.GetString(0));
            Assert.AreEqual("test2", excelReader.GetString(1));
            Assert.AreEqual("test3", excelReader.GetString(2));


            excelReader.Close();



        }

        [TestMethod]
        public void Issue_7433_IllegalOleAutDate()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_7433_IllegalOleAutDate"));

            DataSet dataSet = excelReader.AsDataSet(true);

            Assert.AreEqual(3.10101195608231E+17, dataSet.Tables[0].Rows[0][0]);
            Assert.AreEqual("B221055625", dataSet.Tables[0].Rows[1][0]);
            Assert.AreEqual(4.12721197309241E+17, dataSet.Tables[0].Rows[2][0]);

            excelReader.Close();
        }

        [TestMethod]
        public void Issue_BoolFormula()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_BoolFormula"));


            DataSet dataSet = excelReader.AsDataSet(true);

            Assert.AreEqual(true, dataSet.Tables[0].Rows[0][0]);



            excelReader.Close();
        }

        [TestMethod]
        public void Issue_Decimal_1109_Test()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Decimal_1109"));

            DataSet dataSet = excelReader.AsDataSet(true);

            Assert.AreEqual(3.14159, dataSet.Tables[0].Rows[0][0]);

            const double val1 = -7080.61;
            double val2 = (double)dataSet.Tables[0].Rows[0][1];
            Assert.AreEqual(val1, val2);

            excelReader.Close();
        }

        [TestMethod]
        public void Issue_Encoding_1520_Test()
        {
			IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Encoding_1520"));

            DataSet dataSet = excelReader.AsDataSet();

            string val1 = "Simon Hodgetts";
            string val2 = dataSet.Tables[0].Rows[2][0].ToString();
            Assert.AreEqual(val1, val2);

            val1 = "John test";
            val2 = dataSet.Tables[0].Rows[1][0].ToString();
            Assert.AreEqual(val1, val2);

            //librement réutilisable
            val1 = "librement réutilisable";
            val2 = dataSet.Tables[0].Rows[7][0].ToString();
            Assert.AreEqual(val1, val2);

            val2 = dataSet.Tables[0].Rows[8][0].ToString();
            Assert.AreEqual(val1, val2);

            excelReader.Close();
        }



        [TestMethod]
        public void MultiSheetTest()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("TestMultiSheet"));

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

            excelReader.Close();
        }




        [TestMethod]
        public void Test_num_double_date_bool_string()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_num_double_date_bool_string"));

            DataSet dataSet = excelReader.AsDataSet(true);

            Assert.AreEqual(30, dataSet.Tables[0].Rows.Count);
            Assert.AreEqual(6, dataSet.Tables[0].Columns.Count);

            Assert.AreEqual(1, int.Parse(dataSet.Tables[0].Rows[0][0].ToString()));
            Assert.AreEqual(1346269, int.Parse(dataSet.Tables[0].Rows[29][0].ToString()));

            //bool
            Assert.IsNotNull(dataSet.Tables[0].Rows[22][3].ToString());
            Assert.AreEqual(dataSet.Tables[0].Rows[22][3], true);

            //double + Formula
            Assert.AreEqual(1.02, double.Parse(dataSet.Tables[0].Rows[0][1].ToString()));
            Assert.AreEqual(4.08, double.Parse(dataSet.Tables[0].Rows[2][1].ToString()));
            Assert.AreEqual(547608330.24, double.Parse(dataSet.Tables[0].Rows[29][1].ToString()));

            //Date + Formula
            Assert.AreEqual(new DateTime(2009, 5, 11), dataSet.Tables[0].Rows[0][2]);
            Assert.AreEqual(new DateTime(2009, 11, 30), dataSet.Tables[0].Rows[29][2]);

            //Custom Date Time + Formula
            var s = dataSet.Tables[0].Rows[0][5].ToString();
            Assert.AreEqual(new DateTime(2009, 5, 7, 11, 1, 2), DateTime.Parse(s));
            s = dataSet.Tables[0].Rows[1][5].ToString();
            Assert.AreEqual(new DateTime(2009, 5, 8, 11, 1, 2), DateTime.Parse(s));

            //DBNull value
            Assert.AreEqual(DBNull.Value, dataSet.Tables[0].Rows[1][4]);

            excelReader.Close();
        }

        [TestMethod]
        public void Issue_11479_BlankSheet()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Issue_11479_BlankSheet"));

            //DataSet result = excelReader.AsDataSet(true);

            excelReader.Read();
            Assert.AreEqual(5, excelReader.FieldCount);
            excelReader.NextResult();
            excelReader.Read();
            Assert.AreEqual(0, excelReader.FieldCount);

            excelReader.NextResult();
            excelReader.Read();
            Assert.AreEqual(0, excelReader.FieldCount);

            excelReader.Close();


        }

        [TestMethod]
        public void Test_BlankHeader()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_BlankHeader"));

            excelReader.Read();
            Assert.AreEqual(6, excelReader.FieldCount);
            excelReader.Read();
            for (int i = 0; i < excelReader.FieldCount; i++)
            {
                Console.WriteLine("{0}:{1}", i, excelReader.GetString(i));
            }

            excelReader.Close();
        }

        [TestMethod]
        public void Test_OpenOffice()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_OpenOffice"));
            excelReader.IsFirstRowAsColumnNames = true;

            DoOpenOfficeTest(excelReader);
        }

        /// <summary>
        /// Issue 11 - OpenOffice files were skipping the first row if IsFirstRowAsColumnNames = false;
        /// </summary>
        [TestMethod]
        public void GitIssue_11_OpenOffice_Row_Count()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_OpenOffice"));
            excelReader.IsFirstRowAsColumnNames = false;

            var dataset = excelReader.AsDataSet();
            Assert.AreEqual(34, dataset.Tables[0].Rows.Count);
        }

        /// <summary>
        /// This test is to ensure that we get the same results from an xls saved in excel vs open office
        /// </summary>
        [TestMethod]
        public void Test_OpenOffice_SavedInExcel()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Excel_OpenOffice"));
            excelReader.IsFirstRowAsColumnNames = true;

            DoOpenOfficeTest(excelReader);


        }

		[TestMethod]
		public void Test_Issue_11601_ReadSheetnames()
		{
			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Excel_Dataset"));
			
			Assert.AreEqual("test.csv", excelReader.Name);


			excelReader.NextResult();
			Assert.AreEqual("Sheet2", excelReader.Name);

			excelReader.NextResult();
			Assert.AreEqual("Sheet3", excelReader.Name);

		}

        public static void DoOpenOfficeTest(IExcelDataReader excelReader)
        {
            Assert.IsTrue(excelReader.IsValid);

            excelReader.Read();
            Assert.AreEqual(6, excelReader.FieldCount);
            Assert.AreEqual("column a", excelReader.GetName(0));
            Assert.AreEqual(" column b", excelReader.GetName(1));
            Assert.AreEqual(" column b", excelReader.GetName(2));
            Assert.IsNull(excelReader.GetName(3));
            Assert.AreEqual("column e", excelReader.GetName(4));
            Assert.AreEqual(" column b", excelReader.GetName(5));

            Assert.AreEqual(2, excelReader.GetInt32(0));
            Assert.AreEqual("b", excelReader.GetString(1));
            Assert.AreEqual("c", excelReader.GetString(2));
            Assert.AreEqual("d", excelReader.GetString(3));
            Assert.AreEqual(" e ", excelReader.GetString(4));

            excelReader.Read();
            Assert.AreEqual(6, excelReader.FieldCount);
            Assert.AreEqual(3, excelReader.GetInt32(0));
            Assert.AreEqual(2, excelReader.GetInt32(1));
            Assert.AreEqual(3, excelReader.GetInt32(2));
            Assert.AreEqual(4, excelReader.GetInt32(3));
            Assert.AreEqual(5, excelReader.GetInt32(4));

            excelReader.Read();
            Assert.AreEqual(6, excelReader.FieldCount);
            Assert.AreEqual(4, excelReader.GetInt32(0));
            Assert.AreEqual(new DateTime(2012, 10, 13), excelReader.GetDateTime(1));
            Assert.AreEqual(new DateTime(2012, 10, 14), excelReader.GetDateTime(2));
            Assert.AreEqual(new DateTime(2012, 10, 15), excelReader.GetDateTime(3));
            Assert.AreEqual(new DateTime(2012, 10, 16), excelReader.GetDateTime(4));

            for (int i = 4; i < 34; i++)
            {
                excelReader.Read();
                Assert.AreEqual(i + 1, excelReader.GetInt32(0));
                Assert.AreEqual(i + 2, excelReader.GetInt32(1));
                Assert.AreEqual(i + 3, excelReader.GetInt32(2));
                Assert.AreEqual(i + 4, excelReader.GetInt32(3));
                Assert.AreEqual(i + 5, excelReader.GetInt32(4));
            }

            excelReader.NextResult();
            excelReader.Read();
            Assert.AreEqual(0, excelReader.FieldCount);

            excelReader.NextResult();
            excelReader.Read();
            Assert.AreEqual(0, excelReader.FieldCount);

            //test dataset

            DataSet result = excelReader.AsDataSet(true);
            Assert.AreEqual(1, result.Tables.Count);
            Assert.AreEqual(6, result.Tables[0].Columns.Count);
            Assert.AreEqual(33, result.Tables[0].Rows.Count);

            Assert.AreEqual("column a", result.Tables[0].Columns[0].ColumnName);
            Assert.AreEqual(" column b", result.Tables[0].Columns[1].ColumnName);
            Assert.AreEqual(" column b_1", result.Tables[0].Columns[2].ColumnName);
            Assert.AreEqual("Column3", result.Tables[0].Columns[3].ColumnName);
            Assert.AreEqual("column e", result.Tables[0].Columns[4].ColumnName);
            Assert.AreEqual(" column b_2", result.Tables[0].Columns[5].ColumnName);

            Assert.AreEqual(2, Convert.ToInt32(result.Tables[0].Rows[0][0]));
            Assert.AreEqual("b", result.Tables[0].Rows[0][1]);
            Assert.AreEqual("c", result.Tables[0].Rows[0][2]);
            Assert.AreEqual("d", result.Tables[0].Rows[0][3]);
            Assert.AreEqual(" e ", result.Tables[0].Rows[0][4]);

            Assert.AreEqual(3, Convert.ToInt32(result.Tables[0].Rows[1][0]));
            Assert.AreEqual(2, Convert.ToInt32(result.Tables[0].Rows[1][1]));
            Assert.AreEqual(3, Convert.ToInt32(result.Tables[0].Rows[1][2]));
            Assert.AreEqual(4, Convert.ToInt32(result.Tables[0].Rows[1][3]));
            Assert.AreEqual(5, Convert.ToInt32(result.Tables[0].Rows[1][4]));

            Assert.AreEqual(4, Convert.ToInt32(result.Tables[0].Rows[2][0]));
            Assert.AreEqual(new DateTime(2012, 10, 13), result.Tables[0].Rows[2][1]);
            Assert.AreEqual(new DateTime(2012, 10, 14), result.Tables[0].Rows[2][2]);
            Assert.AreEqual(new DateTime(2012, 10, 15), result.Tables[0].Rows[2][3]);
            Assert.AreEqual(new DateTime(2012, 10, 16), result.Tables[0].Rows[2][4]);

            for (int i = 4; i < 33; i++)
            {
                Assert.AreEqual(i + 2, Convert.ToInt32(result.Tables[0].Rows[i][0]));
                Assert.AreEqual(i + 3, Convert.ToInt32(result.Tables[0].Rows[i][1]));
                Assert.AreEqual(i + 4, Convert.ToInt32(result.Tables[0].Rows[i][2]));
                Assert.AreEqual(i + 5, Convert.ToInt32(result.Tables[0].Rows[i][3]));
                Assert.AreEqual(i + 6, Convert.ToInt32(result.Tables[0].Rows[i][4]));
            }
            excelReader.Close();
        }

        [TestMethod]
        public void UnicodeCharsTest()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("TestUnicodeChars"));

            DataTable result = excelReader.AsDataSet().Tables[0];

            Assert.AreEqual(3, result.Rows.Count);
            Assert.AreEqual(8, result.Columns.Count);
            Assert.AreEqual("\u00e9\u0417", result.Rows[1][0].ToString());

            excelReader.Close();
        }

        [TestMethod]
        public void UncalculatedTest()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Uncalculated"));

            var dataset = excelReader.AsDataSet();
            Assert.IsNotNull(dataset);
            Assert.AreNotEqual(dataset.Tables.Count, 0);
            var table = dataset.Tables[0];
            Assert.IsNotNull(table);

            Assert.AreEqual("1", table.Rows[1][0].ToString());
            Assert.AreEqual("3", table.Rows[1][2].ToString());
            Assert.AreEqual("3", table.Rows[1][4].ToString());

            excelReader.Close();
        }

		[TestMethod]
		public void Issue_11570_Excel2013()
		{
			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_11570_Excel2013"));

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

			excelReader.Close();
		}
		
		[TestMethod]
		//"Issue will mot be resolved as codepage 27651 is not supported in .net \"System.NotSupportedException : No data is available for encoding 27651.\"")]
		[Ignore("codepage 27651 is not supported in .net")]
		public void Issue_11572_CodePage()
		{
			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_11572_CodePage"));

			var dataset = excelReader.AsDataSet();

			

			excelReader.Close();
		}		
		
        /// <summary>
        /// Not fixed yet
        /// </summary>
		[TestMethod]
		public void Issue_11545_NoIndex()
		{
            Assert.Inconclusive("not fixed yet");
			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_11545_NoIndex"));
			excelReader.IsFirstRowAsColumnNames = true;
			var dataset = excelReader.AsDataSet();

			Assert.AreEqual("CI2229         ", dataset.Tables[0].Rows[0][0]);
			Assert.AreEqual("12069E01018A1  ", dataset.Tables[0].Rows[0][6]);
			Assert.AreEqual(new DateTime(2012, 03, 01), dataset.Tables[0].Rows[0][8]);
			excelReader.Close();
		}


		[TestMethod]
		public void Issue_11573_BlankValues()
		{
            
			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_11573_BlankValues"));
			excelReader.IsFirstRowAsColumnNames = false;
			var dataset = excelReader.AsDataSet();

			Assert.AreEqual(1D, dataset.Tables[0].Rows[12][0]);
			Assert.AreEqual("070202", dataset.Tables[0].Rows[12][1]);

			excelReader.Close();
		}
		
		[TestMethod]
		public void Issue_DateFormatButNotDate()
		{
			//we want to make sure that if a cell is formatted as a date but it's contents are not a date then
			//the output is not a date
			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_DateFormatButNotDate"), true);

			excelReader.Read();
			Assert.AreEqual("columna", excelReader.GetValue(0) );
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
			excelReader.Close();
		}

		[TestMethod]
		public void Issue_11642_ValuesNotLoaded()
		{
			//Excel.Log.Log.InitializeWith<Log4NetLog>();

			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_11642_ValuesNotLoaded"));
			excelReader.IsFirstRowAsColumnNames = false;
			var dataset = excelReader.AsDataSet();

			Assert.AreEqual("431113*", dataset.Tables[2].Rows[29][1].ToString());
			Assert.AreEqual("024807", dataset.Tables[2].Rows[36][1].ToString());
			Assert.AreEqual("160019", dataset.Tables[2].Rows[53][1].ToString());

			excelReader.Close();
		}

		[TestMethod]
		public void Issue_11636_BiffStream()
		{
			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_11636_BiffStream"), ReadOption.Loose);
			excelReader.IsFirstRowAsColumnNames = false;
			var dataset = excelReader.AsDataSet();

			//check a couple of values
			Assert.AreEqual("SP011", dataset.Tables[0].Rows[9][0]);
			Assert.AreEqual(9.9, dataset.Tables[0].Rows[32][11]);
			Assert.AreEqual(78624.44, dataset.Tables[1].Rows[27][12]);

			excelReader.Close();
		}

        /// <summary>
        /// Not fixed yet
        /// The problem occurs with unseekable stream and loigc related to minifat that uses seek
        /// It should probably only use seek if it needs to go backwards, I think at the moment it uses seek all the time
        /// which is probably not good for performance
        /// </summary>
		[TestMethod]
		public void Issue_11639_11644_ForwardOnlyStream()
		{
            Assert.Inconclusive("Not fixed yet");
			//Excel.Log.Log.InitializeWith<Log4NetLog>();
			var forwardStream = SeekErrorMemoryStream.CreateFromStream(Helper.GetTestWorkbook("Test_OpenOffice"));

			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateBinaryReader(forwardStream);
			excelReader.IsFirstRowAsColumnNames = false;
			var dataset = excelReader.AsDataSet();


			excelReader.Close();
		}

        /// <summary>
        /// Not fixed yet
        /// The problem occurs with unseekable stream and loigc related to minifat that uses seek
        /// It should probably only use seek if it needs to go backwards, I think at the moment it uses seek all the time
        /// which is probably not good for performance
        /// </summary>
		[TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void Issue_12556_corrupt()
		{
            //Excel.Log.Log.InitializeWith<Log4NetLog>();
            var forwardStream = Helper.GetTestWorkbook("Test_Issue_12556_corrupt");

			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateBinaryReader(forwardStream);
			excelReader.IsFirstRowAsColumnNames = false;
			var dataset = excelReader.AsDataSet();


			excelReader.Close();
		}

        /// <summary>
        /// Some spreadsheets were crashing with index out of range error (from SSRS)
        /// </summary>
        [TestMethod]
        public void Test_Issue_11818_OutOfRange()
		{
#if !NETCOREAPP1_0
			ExcelDataReader.Log.Log.InitializeWith<Log.Logger.Log4NetLog>();
#endif
			IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Issue_11818_OutOfRange"), ReadOption.Loose);
            excelReader.IsFirstRowAsColumnNames = false;
            var dataset = excelReader.AsDataSet();

            Assert.AreEqual("Total Revenue", dataset.Tables[0].Rows[10][0]);

            excelReader.Close();
		}

        [TestMethod]
        public void Test_Issue_111_NoRowRecords()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_git_issue_111_NoRowRecords"), ReadOption.Loose))
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
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_Git_Issue_145"), ReadOption.Loose);

            excelReader.Read();
            excelReader.Read();
            excelReader.Read();

            string value = excelReader.GetString(3);

            Assert.AreEqual("Japanese Government Bonds held by the Bank of Japan", value);
        }

        [TestMethod]
        public void Test_GitIssue_152_SheetName_UTF16LE_Compressed()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_git_issue_152"));
            var dataset = excelReader.AsDataSet();

            Assert.AreEqual("åäöñ", dataset.Tables[0].TableName);

            excelReader.Close();
        }

        [TestMethod]
        public void Test_GitIssue_152_Cell_UTF16LE_Compressed()
        {
            IExcelDataReader excelReader =
                ExcelReaderFactory.CreateBinaryReader(Helper.GetTestWorkbook("Test_git_issue_152"));
            var dataset = excelReader.AsDataSet();

            Assert.AreEqual("åäöñ", dataset.Tables[0].Rows[0][0]);

            excelReader.Close();
        }
    }
}