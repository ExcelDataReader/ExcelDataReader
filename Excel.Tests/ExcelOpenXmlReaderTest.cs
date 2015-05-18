using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Threading;
using Excel;
using Excel.Tests.Log.Logger;
#if LEGACY
#else
using ExcelDataReader.Portable.Core;
using ExcelDataReader.Portable.IO;
using Moq;
using PCLStorage;
#endif

using Microsoft.VisualStudio.TestTools.UnitTesting;


#if LEGACY
namespace Excel.Tests
#else
namespace ExcelDataReader.Tests
#endif
{
	[TestClass]
	public class ExcelOpenXmlReaderTest
	{
		[TestInitialize]
		public void TestInitialize()
		{
			//Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE", false);
		}

		[TestMethod]
		public void GitIssue_29_ReadSheetStatesReadsCorrectly()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Excel_Dataset"));

			Assert.AreEqual("hidden", excelReader.VisibleState);

			excelReader.NextResult();
			Assert.AreEqual("visible", excelReader.VisibleState);

			excelReader.NextResult();
			Assert.AreEqual("veryhidden", excelReader.VisibleState);
		}

		[TestMethod]
		public void GitIssue_29_AsDataSetProvidesCorrectSheetState()
		{
			IExcelDataReader reader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Excel_Dataset"));

			var dataset = reader.AsDataSet();

			reader.Close();

			Assert.IsTrue(dataset != null);
			Assert.AreEqual(3, dataset.Tables.Count);
			Assert.AreEqual("hidden", dataset.Tables[0].ExtendedProperties["visiblestate"]);
			Assert.AreEqual("visible", dataset.Tables[1].ExtendedProperties["visiblestate"]);
			Assert.AreEqual("veryhidden", dataset.Tables[2].ExtendedProperties["visiblestate"]);
		}

		[TestMethod]
		public void Issue_11516_workbook_with_single_sheet_should_not_return_empty_dataset()
		{
			IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_11516_Single_Tab"));

			Assert.AreEqual(1, reader.ResultsCount);

			DataSet dataset = reader.AsDataSet();

			reader.Close();

			Assert.IsTrue(dataset != null);
			Assert.AreEqual(1, dataset.Tables.Count);
			Assert.AreEqual(260, dataset.Tables[0].Rows.Count);
			Assert.AreEqual(29, dataset.Tables[0].Columns.Count);
		}

		[TestMethod]
		public void AsDataset_Test()
		{
			IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTestOpenXml"));

			Assert.AreEqual(3, reader.ResultsCount);

			DataSet dataset = reader.AsDataSet();

			reader.Close();

			Assert.IsTrue(dataset != null);
			Assert.AreEqual(3, dataset.Tables.Count);
			Assert.AreEqual(7, dataset.Tables["Sheet1"].Rows.Count);
			Assert.AreEqual(11, dataset.Tables["Sheet1"].Columns.Count);
		}


		[TestMethod]
		public void ChessTest()
		{
			IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTestChess"));

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
			IExcelDataReader r = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTestMultiSheet"));

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
			table.TableName = r.Name;

			while (r.Read())
			{
				fieldCount = r.FieldCount;
				table.Rows.Add(r.GetInt32(0), r.GetInt32(1), r.GetInt32(2), r.GetInt32(3));
			}

			Assert.AreEqual(12, table.Rows.Count);
			Assert.AreEqual(4, fieldCount);
			Assert.AreEqual(2, table.Rows[11][3]);


			r.NextResult();
			table.TableName = r.Name;
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
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_num_double_date_bool_string"));

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
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest10x10000"));

			DataTable result = excelReader.AsDataSet().Tables[0];

			excelReader.Close();

			Assert.AreEqual(10000, result.Rows.Count);
			Assert.AreEqual(10, result.Columns.Count);
			Assert.AreEqual("1x2", result.Rows[1][1]);
			Assert.AreEqual("1x10", result.Rows[1][9]);
			Assert.AreEqual("1x1", result.Rows[9999][0]);
			Assert.AreEqual("1x10", result.Rows[9999][9]);
		}

		[TestMethod]
		public void Dimension10x10Test()
		{
			IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest10x10"));

			DataTable result = excelReader.AsDataSet().Tables[0];

			excelReader.Close();

			Assert.AreEqual(10, result.Rows.Count);
			Assert.AreEqual(10, result.Columns.Count);
			Assert.AreEqual("10x10", result.Rows[1][0]);
			Assert.AreEqual("10x27", result.Rows[9][9]);
		}

		[TestMethod]
		public void Dimension255x10Test()
		{
			IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest255x10"));


			DataTable result = excelReader.AsDataSet().Tables[0];

			excelReader.Close();

			Assert.AreEqual(10, result.Rows.Count);
			Assert.AreEqual(255, result.Columns.Count);
			Assert.AreEqual("1", result.Rows[9][254].ToString());
			Assert.AreEqual("one", result.Rows[1][1].ToString());
		}

		[TestMethod]
		public void DoublePrecisionTest()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTestDoublePrecision"));

			DataTable result = excelReader.AsDataSet().Tables[0];

			excelReader.Close();

			Assert.AreEqual(10, result.Rows.Count);

			const double excelPI = 3.14159265358979;

			Assert.AreEqual(+excelPI, result.Rows[2][1]);
			Assert.AreEqual(-excelPI, result.Rows[3][1]);

			Assert.AreEqual(+excelPI * 1.0e-300, result.Rows[4][1]);
			Assert.AreEqual(-excelPI * 1.0e-300, result.Rows[5][1]);

			Assert.AreEqual(+excelPI * 1.0e300, (double)result.Rows[6][1], 1e286); //only accurate to 1e286 because excel only has 15 digits precision
			Assert.AreEqual(-excelPI * 1.0e300, (double)result.Rows[7][1], 1e286);

			Assert.AreEqual(+excelPI * 1.0e14, result.Rows[8][1]);
			Assert.AreEqual(-excelPI * 1.0e14, result.Rows[9][1]);
		}

		[TestMethod]
		public void Fail_Test()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("TestFail_Binary"));

			Assert.AreEqual(false, excelReader.IsValid);
			Assert.AreEqual(true, excelReader.IsClosed);
#if LEGACY
			Assert.AreEqual("Cannot find central directory", excelReader.ExceptionMessage);

#else
            Assert.AreEqual("End of Central Directory record could not be found.", excelReader.ExceptionMessage);        
#endif
		}


		[TestMethod]
		public void Issue_Date_and_Time_1468_Test()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Encoding_1520"));

			DataSet dataSet = excelReader.AsDataSet(true);

			excelReader.Close();

			string val1 = new DateTime(2009, 05, 01).ToShortDateString();
			string val2 = DateTime.Parse(dataSet.Tables[0].Rows[1][1].ToString()).ToShortDateString();

			Assert.AreEqual(val1, val2);

			val1 = new DateTime(2009, 1, 1, 11, 0, 0).ToShortTimeString();
			val2 = DateTime.Parse(dataSet.Tables[0].Rows[2][4].ToString()).ToShortTimeString();

			Assert.AreEqual(val1, val2);
		}

		[TestMethod]
		public void Issue_8536_Test()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_8536"));

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
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_11397"));

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
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_4031_NullColumn"));

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
		public void Issue_4145()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_4145"));

			excelReader.IsFirstRowAsColumnNames = true;
			DataSet result = excelReader.AsDataSet();


			while (excelReader.Read())
			{

			}


			excelReader.Close();



		}

		[TestMethod]
		public void Issue_10725()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_10725"));


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
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_11435_Colors"));


			excelReader.Read();

			Assert.AreEqual("test1", excelReader.GetString(0));
			Assert.AreEqual("test2", excelReader.GetString(1));
			Assert.AreEqual("test3", excelReader.GetString(2));

			DataSet dataSet = excelReader.AsDataSet(true);

			Assert.AreEqual("test1", dataSet.Tables[0].Rows[0][0].ToString());
			Assert.AreEqual("test2", dataSet.Tables[0].Rows[0][1].ToString());
			Assert.AreEqual("test3", dataSet.Tables[0].Rows[0][2].ToString());




			excelReader.Close();



		}

		[TestMethod]
		public void Issue_7433_IllegalOleAutDate()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_7433_IllegalOleAutDate"));

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
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_BoolFormula"));


			DataSet dataSet = excelReader.AsDataSet(true);

			Assert.AreEqual(true, dataSet.Tables[0].Rows[0][0]);



			excelReader.Close();
		}

		[TestMethod]
		public void Issue_Decimal_1109_Test()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Decimal_1109"));

			DataSet dataSet = excelReader.AsDataSet(true);

			excelReader.Close();


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
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Encoding_1520"));

			DataSet dataSet = excelReader.AsDataSet();

			excelReader.Close();

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
		}

		[TestMethod]
		public void Issue_FileLock_5161()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTestMultiSheet"));


			//read something from the 3rd sheet
			int i = 0;
			do
			{
				if (i == 0)
					excelReader.Read();
			} while (excelReader.NextResult());


			//bug was exposed here 
			excelReader.Close();
		}

		[TestMethod]
		public void Test_BlankHeader()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_BlankHeader"));

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
		public void Test_OpenOffice_SavedInExcel()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Excel_OpenOffice"));
			excelReader.IsFirstRowAsColumnNames = true;

			ExcelBinaryReaderTest.DoOpenOfficeTest(excelReader);


		}

		[TestMethod]
		public void Test_Issue_11601_ReadSheetnames()
		{
			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Excel_Dataset"));

			Assert.AreEqual("test.csv", excelReader.Name);


			excelReader.NextResult();
			Assert.AreEqual("Sheet2", excelReader.Name);

			excelReader.NextResult();
			Assert.AreEqual("Sheet3", excelReader.Name);

		}

		[TestMethod]
		public void MultiSheetTest()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTestMultiSheet"));

			DataSet result = excelReader.AsDataSet();

			excelReader.Close();

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


		[TestMethod]
		public void Test_num_double_date_bool_string()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_num_double_date_bool_string"));

			DataSet dataSet = excelReader.AsDataSet(true);

			excelReader.Close();

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
			string s = dataSet.Tables[0].Rows[0][2].ToString();
			Assert.AreEqual(new DateTime(2009, 5, 11), DateTime.Parse(s));
			s = dataSet.Tables[0].Rows[29][2].ToString();
			Assert.AreEqual(new DateTime(2009, 11, 30), DateTime.Parse(s));

			//Custom Date Time + Formula
			s = dataSet.Tables[0].Rows[0][5].ToString();
			Assert.AreEqual(new DateTime(2009, 5, 7, 11, 1, 2), DateTime.Parse(s));
			s = dataSet.Tables[0].Rows[1][5].ToString();
			Assert.AreEqual(new DateTime(2009, 5, 8, 11, 1, 2), DateTime.Parse(s));

			//DBNull value
			Assert.AreEqual(DBNull.Value, dataSet.Tables[0].Rows[1][4]);
		}

		[TestMethod]
		public void Issue_11479_BlankSheet()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xIssue_11479_BlankSheet"));

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
		public void Issue_11522_OpenXml()
		{
			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_11522_OpenXml"));
			excelReader.IsFirstRowAsColumnNames = true;
			DataSet result = excelReader.AsDataSet(true);

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
			excelReader.Close();


		}

		[TestMethod]
		public void UnicodeCharsTest()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTestUnicodeChars"));

			DataTable result = excelReader.AsDataSet().Tables[0];

			excelReader.Close();

			Assert.AreEqual(3, result.Rows.Count);
			Assert.AreEqual(8, result.Columns.Count);
			Assert.AreEqual(Helper.GetKey("TestUnicodePos2x1"), result.Rows[1][0].ToString());
		}

#if !LEGACY
        [TestMethod]
        public void ZipWorker_Extract_Test()
        {
            var zipper = new ZipWorker(FileSystem.Current, new FileHelper());

            //this first one isn't a valid xlsx so we are expecting no side effects in the directory tree
            zipper.Extract(Helper.GetTestWorkbook("TestChess"));
            Assert.AreEqual(false, Directory.Exists(zipper.TempPath));
            Assert.AreEqual(false, zipper.IsValid);

            //this one is valid so we expect to find the files
            zipper.Extract(Helper.GetTestWorkbook("xTestOpenXml"));

            Assert.AreEqual(true, Directory.Exists(zipper.TempPath));
            Assert.AreEqual(true, zipper.IsValid);

            string tPath = zipper.TempPath;

            //make sure that dispose gets rid of the files
            zipper.Dispose();

            Assert.AreEqual(false, Directory.Exists(tPath));
        }

        private class FileHelper : IFileHelper
        {
            public string GetTempPath()
            {
                return System.IO.Path.GetTempPath();
            }
        }
#endif


		[TestMethod]
		public void Issue_DateFormatButNotDate()
		{
			//we want to make sure that if a cell is formatted as a date but it's contents are not a date then
			//the output is not a date (it was ending up as datetime.min)
			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_DateFormatButNotDate"));

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
			excelReader.Close();
		}

		[TestMethod]
		public void Issue_11573_BlankValues()
		{
#if !LEGACY
			ExcelDataReader.Portable.Log.Log.InitializeWith<Log4NetLog>();
#endif
			IExcelDataReader excelReader =
	ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_11573_BlankValues"));
			excelReader.IsFirstRowAsColumnNames = false;
			var dataset = excelReader.AsDataSet();

			Assert.AreEqual(1D, dataset.Tables[0].Rows[12][0]);
			Assert.AreEqual(070202D, dataset.Tables[0].Rows[12][1]);

			excelReader.Close();
		}

		[TestMethod]
		public void Issue_11773_Exponential()
		{
			IExcelDataReader excelReader =
				ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_11773_Exponential"));
			excelReader.IsFirstRowAsColumnNames = true;
			var dataset = excelReader.AsDataSet();

			Assert.AreEqual(2566.37168141593D, double.Parse(dataset.Tables[0].Rows[0][6].ToString()));

			excelReader.Close();
		}

		[TestMethod]
		public void Issue_11773_Exponential_Commas()
		{
			Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE", false);

			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_11773_Exponential"));
			excelReader.IsFirstRowAsColumnNames = true;
			var dataset = excelReader.AsDataSet();

			Assert.AreEqual(2566.37168141593D, double.Parse(dataset.Tables[0].Rows[0][6].ToString()));

			excelReader.Close();
		}

		[TestMethod]
		public void Test_googlesourced()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_googlesourced"));
			excelReader.IsFirstRowAsColumnNames = true;
			var dataset = excelReader.AsDataSet();

			Assert.AreEqual("9583638582", dataset.Tables[0].Rows[0][0].ToString());
			Assert.AreEqual(4, dataset.Tables[0].Rows.Count);
			Assert.AreEqual(6, dataset.Tables[0].Columns.Count);

			excelReader.Close();
		}

		[TestMethod]
		public void Test_Issue_12667_GoogleExport_MissingColumns()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_Issue_12667_GoogleExport_MissingColumns"));
			excelReader.IsFirstRowAsColumnNames = true;
			var dataset = excelReader.AsDataSet();

			Assert.AreEqual(7, dataset.Tables[0].Columns.Count); //6 with data + 1 that is present but no data in it
			Assert.AreEqual(19, dataset.Tables[0].Rows.Count);

			excelReader.Close();
		}

		/// <summary>
		/// Makes sure that we can read data from the first roiw of last sheet
		/// </summary>
		[TestMethod]
		public void Issue_12271_NextResultSet()
		{
			IExcelDataReader excelReader =
					ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_LotsOfSheets"));
			//excelReader.IsFirstRowAsColumnNames = true;

			do
			{
				excelReader.Read();

				if (excelReader.FieldCount == 0)
					continue;

				//ignore sheets beginning with $e
				if (excelReader.Name.StartsWith("$e", StringComparison.InvariantCultureIgnoreCase))
					continue;

				Assert.AreEqual("StaffName", excelReader.GetString(0));
			} while (excelReader.NextResult());

			excelReader.Close();
		}

		[TestMethod]
		public void Test_XLSM_Support()
		{
			IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTest_ExcelMacroEnabledWorkbook"));
			DataSet ds = excelReader.AsDataSet();
			Assert.IsNotNull(ds, "Expected that DataSet can be created from *.xlsm type file");
		}
	}
}