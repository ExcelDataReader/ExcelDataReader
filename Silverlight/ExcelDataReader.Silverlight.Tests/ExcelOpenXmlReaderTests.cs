namespace ExcelDataReader.Silverlight.Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Windows;
    using System.Windows.Resources;
    using ExcelDataReader.Silverlight.Data;
    using ExcelDataReader.Silverlight.Data.Example;
    using Microsoft.Silverlight.Testing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class ExcelOpenXmlReaderTests
    {
        [TestMethod]
        [Tag("Xml")]
        public void AsDataSet_ValidFile_ReturnsPopulatedDataSet()
        {
            StreamResourceInfo sri = Application.GetResourceStream(new Uri("Resources/OpenXmlFile.xlsx", UriKind.Relative));
            Stream fileStream = sri.Stream;

            var reader = ExcelReaderFactory.CreateOpenXmlReader(fileStream) as ExcelOpenXmlReader;
            reader.WorkBookFactory = new WorkBookFactory();

            var dataSet = reader.AsWorkBook();

            Assert.AreEqual(dataSet.WorkSheets.Count, 2);

            var firstTable = dataSet.WorkSheets.First();
            Assert.AreEqual(firstTable.Columns.Count, 2);
            Assert.AreEqual(firstTable.Rows.Count, 2);
            var firstRowValues = firstTable.Rows.ElementAt(0).Values.OfType<object>();
            Assert.AreEqual(firstRowValues.ElementAt(0).ToString(), "Sheet1.A1");
            Assert.AreEqual(firstRowValues.ElementAt(1).ToString(), "Sheet1.B1");
            var secondRowValues = firstTable.Rows.ElementAt(1).Values.OfType<object>();
            Assert.AreEqual(secondRowValues.ElementAt(0).ToString(), "Sheet1.A2");
            Assert.AreEqual(secondRowValues.ElementAt(1).ToString(), "Sheet1.B2");

            var secondTable = dataSet.WorkSheets.ElementAt(1);
            Assert.AreEqual(secondTable.Columns.Count, 2);
            Assert.AreEqual(secondTable.Rows.Count, 2);
            firstRowValues = secondTable.Rows.ElementAt(0).Values.OfType<object>();
            Assert.AreEqual(firstRowValues.ElementAt(0).ToString(), "Sheet2.A1");
            Assert.AreEqual(firstRowValues.ElementAt(1).ToString(), "Sheet2.B1");
            secondRowValues = secondTable.Rows.ElementAt(1).Values.OfType<object>();
            Assert.AreEqual(secondRowValues.ElementAt(0).ToString(), "Sheet2.A2");
            Assert.AreEqual(secondRowValues.ElementAt(1).ToString(), "Sheet2.B2");
        }

        //private static ExcelOpenXmlReader GetReader()
        //{
        //    StreamResourceInfo sri = Application.GetResourceStream(new Uri("Resources/OpenXmlFile.xlsx", UriKind.Relative));
        //    Stream fileStream = sri.Stream;

        //    var openXmlReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream) as ExcelOpenXmlReader;
        //    openXmlReader.WorkBookFactory = new WorkBookFactory();

        //    return openXmlReader;
        //}

        public static double ParseDouble(string s)
        {
            return double.Parse(s, System.Globalization.CultureInfo.InvariantCulture);
        }

        private static Stream GetTestWorkBook(string workBook)
        {
            StreamResourceInfo sri = Application.GetResourceStream(new Uri(workBook, UriKind.Relative));
            Stream fileStream = sri.Stream;
            return fileStream;
        }

        [TestMethod]
        [Tag("Xml")]
        public void AsDataSet_Test()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/TestChess.xlsx"));
            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            Assert.IsTrue(excelWorkBook != null);
            Assert.AreEqual(1, workSheets.Count);
            Assert.AreEqual(4, workSheets[0].Rows.Count);
            Assert.AreEqual(6, workSheets[0].Columns.Count);

            excelReader.Close();
        }

        [TestMethod]
        [Tag("Xml")]
        public void Fail_Test()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/TestFail_Binary.xls"));

            Assert.AreEqual(false, excelReader.IsValid);
            //Assert.AreEqual(true, excelReader.IsClosed);
            Assert.AreEqual("Cannot find central directory", excelReader.ExceptionMessage);
        }


        [TestMethod]
        [Tag("Xml")]
        public void ChessTest()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/TestChess.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;


            Assert.AreEqual(4, workSheets[0].Rows.Count);
            Assert.AreEqual(6, workSheets[0].Columns.Count);
            Assert.AreEqual("1", workSheets[0].Rows[3].Values[5].ToString());
            Assert.AreEqual("1", workSheets[0].Rows[2].Values[0].ToString());

            excelReader.Close();
        }


        [TestMethod]
        [Tag("Xml")]
        public void Dimension10x10000Test()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/Test10x10000.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            Assert.AreEqual(10000, workSheets[0].Rows.Count);
            Assert.AreEqual(10, workSheets[0].Columns.Count);
            Assert.AreEqual("1x2", workSheets[0].Rows[1].Values[1]);
            Assert.AreEqual("1x10", workSheets[0].Rows[1].Values[9]);
            Assert.AreEqual("1x1", workSheets[0].Rows[9999].Values[0]);
            Assert.AreEqual("1x10", workSheets[0].Rows[9999].Values[9]);

            excelReader.Close();
        }


        [TestMethod]
        [Tag("Xml")]
        public void Dimension10x10Test()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/Test10x10.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            Assert.AreEqual(10, workSheets[0].Rows.Count);
            Assert.AreEqual(10, workSheets[0].Columns.Count);
            Assert.AreEqual("10x10", workSheets[0].Rows[1].Values[0]);
            Assert.AreEqual("10x27", workSheets[0].Rows[9].Values[9]);

            excelReader.Close();
        }


        [TestMethod]
        [Tag("Xml")]
        public void Dimension255x10Test()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/Test255x10.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            Assert.AreEqual(10, workSheets[0].Rows.Count);
            Assert.AreEqual(255, workSheets[0].Columns.Count);
            Assert.AreEqual("1", workSheets[0].Rows[9].Values[254].ToString());
            Assert.AreEqual("one", workSheets[0].Rows[1].Values[1].ToString());

            excelReader.Close();
        }


        [TestMethod]
        [Tag("Xml")]
        public void MultiSheetTest()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/TestMultiSheet.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            Assert.AreEqual(3, workSheets.Count);

            Assert.AreEqual(4, workSheets.FirstOrDefault(ws => ws.Name == "Sheet1").Columns.Count);
            Assert.AreEqual(12, workSheets.FirstOrDefault(ws => ws.Name == "Sheet1").Rows.Count);
            Assert.AreEqual(4, workSheets.FirstOrDefault(ws => ws.Name == "Sheet2").Columns.Count);
            Assert.AreEqual(12, workSheets.FirstOrDefault(ws => ws.Name == "Sheet2").Rows.Count);
            Assert.AreEqual(2, workSheets.FirstOrDefault(ws => ws.Name == "Sheet3").Columns.Count);
            Assert.AreEqual(5, workSheets.FirstOrDefault(ws => ws.Name == "Sheet3").Rows.Count);

            var wks2 = workSheets.FirstOrDefault(ws => ws.Name == "Sheet2");
            Assert.AreEqual("1", wks2.Rows[11].Values[0].ToString());

            var wks1 = workSheets.FirstOrDefault(ws => ws.Name == "Sheet1");
            Assert.AreEqual("2", wks1.Rows[11].Values[3].ToString());

            var wks3 = workSheets.FirstOrDefault(ws => ws.Name == "Sheet3");
            Assert.AreEqual("3", wks3.Rows[4].Values[1].ToString());

            excelReader.Close();
        }


        [TestMethod]
        [Tag("Xml")]
        public void UnicodeCharsTest()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/TestUnicodeChars.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            Assert.AreEqual(3, workSheets[0].Rows.Count);
            Assert.AreEqual(8, workSheets[0].Columns.Count);
            //missing
            excelReader.Close();
        }


        [TestMethod]
        [Tag("Xml")]
        public void DoublePrecisionTest()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/TestDoublePrecision.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            double excelPI = 3.1415926535897900;

            Assert.AreEqual(+excelPI, ParseDouble(workSheets[0].Rows[2].Values[1].ToString()), 1e-14);
            Assert.AreEqual(-excelPI, ParseDouble(workSheets[0].Rows[3].Values[1].ToString()), 1e-14);

            Assert.AreEqual(+excelPI * 1.0e-300, ParseDouble(workSheets[0].Rows[4].Values[1].ToString()), 3e-315);
            Assert.AreEqual(-excelPI * 1.0e-300, ParseDouble(workSheets[0].Rows[5].Values[1].ToString()), 3e-315);

            Assert.AreEqual(+excelPI * 1.0e300, ParseDouble(workSheets[0].Rows[6].Values[1].ToString()), 3e+285);
            Assert.AreEqual(-excelPI * 1.0e300, ParseDouble(workSheets[0].Rows[7].Values[1].ToString()), 3e+287);

            Assert.AreEqual(+excelPI * 1.0e15, ParseDouble(workSheets[0].Rows[8].Values[1].ToString()), 3e+1);
            Assert.AreEqual(-excelPI * 1.0e15, ParseDouble(workSheets[0].Rows[9].Values[1].ToString()), 3e+1);

            excelReader.Close();
        }


        [TestMethod]
        [Tag("Xml")]
        public void Issue_Encoding_1520_Test()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/Test_Encoding_Formula_Date_1520.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;


            string val1 = "Simon Hodgetts";
            string val2 = workSheets[0].Rows[2].Values[0].ToString();
            Assert.AreEqual(val1, val2);

            val1 = "John test";
            val2 = workSheets[0].Rows[1].Values[0].ToString();
            Assert.AreEqual(val1, val2);

            //librement réutilisable
            val1 = "librement réutilisable";
            val2 = workSheets[0].Rows[7].Values[0].ToString();
            Assert.AreEqual(val1, val2);

            val2 = workSheets[0].Rows[8].Values[0].ToString();
            Assert.AreEqual(val1, val2);

            excelReader.Close();
        }


        [TestMethod]
        [Tag("Xml")]
        public void Issue_Date_and_Time_1468_Test()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/Test_Encoding_Formula_Date_1520.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            string val1 = new DateTime(2009, 05, 01).ToShortDateString();
            string val2 = DateTime.Parse(workSheets[0].Rows[1].Values[1].ToString()).ToShortDateString();

            Assert.AreEqual(val1, val2);

            val1 = DateTime.Parse("11:00:00").ToShortTimeString();
            val2 = DateTime.Parse(workSheets[0].Rows[2].Values[4].ToString()).ToShortTimeString();

            Assert.AreEqual(val1, val2);

            excelReader.Close();
        }


        [TestMethod]
        [Tag("Xml")]
        public void Issue_Decimal_1109_Test()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/Test_Decimal_1109.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            Assert.AreEqual(Double.Parse("3.14159"), Double.Parse(workSheets[0].Rows[0].Values[0].ToString()));

            double val1 = -7080.61;
            double val2 = Double.Parse(workSheets[0].Rows[0].Values[1].ToString());
            Assert.AreEqual(val1, val2);

            excelReader.Close();
        }


        [TestMethod]
        [Tag("Xml")]
        public void Issue_EmptyCells_5320_Test()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/TestEmptyCells_5320.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            excelReader.Close();

            string val1 = new DateTime(2009, 1, 31).ToShortDateString();
            string val2Unparsed = workSheets[0].Rows[4].Values[1].ToString();
            Assert.IsFalse(string.IsNullOrEmpty(val2Unparsed));

            string val2 = DateTime.Parse(val2Unparsed).ToShortDateString();
            Assert.AreEqual(val1, val2);

            val1 = new DateTime(2009, 2, 28).ToShortDateString();
            val2Unparsed = workSheets[0].Rows[4].Values[7].ToString();
            Assert.IsFalse(string.IsNullOrEmpty(val2Unparsed));

            val2 = DateTime.Parse(val2Unparsed).ToShortDateString();
            Assert.AreEqual(val1, val2);
        }


        [TestMethod]
        [Tag("Xml")]
        public void Test_num_double_date_bool_string()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/Test_num_double_date_bool_string.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            Assert.AreEqual(30, workSheets[0].Rows.Count);
            Assert.AreEqual(6, workSheets[0].Columns.Count);

            Assert.AreEqual(1, int.Parse(workSheets[0].Rows[0].Values[0].ToString()));
            Assert.AreEqual(1346269, int.Parse(workSheets[0].Rows[29].Values[0].ToString()));

            //double + Formula
            Assert.AreEqual(1.02, double.Parse(workSheets[0].Rows[0].Values[1].ToString()));
            Assert.AreEqual(4.08, double.Parse(workSheets[0].Rows[2].Values[1].ToString()));
            Assert.AreEqual(547608330.24, double.Parse(workSheets[0].Rows[29].Values[1].ToString()));

            //Date + Formula
            Assert.AreEqual(new DateTime(2009, 5, 11).ToShortDateString(), DateTime.Parse(workSheets[0].Rows[0].Values[2].ToString()).ToShortDateString());
            Assert.AreEqual(new DateTime(2009, 11, 30).ToShortDateString(), DateTime.Parse(workSheets[0].Rows[29].Values[2].ToString()).ToShortDateString());

            //Custom Date Time + Formula
            Assert.AreEqual(new DateTime(2009, 5, 7).ToShortDateString(), DateTime.Parse(workSheets[0].Rows[0].Values[5].ToString()).ToShortDateString());
            Assert.AreEqual(new DateTime(2009, 5, 8, 11, 1, 2), DateTime.Parse(workSheets[0].Rows[1].Values[5].ToString()));

            //DBNull value (dont have dbnull when not using datatable
            Assert.IsNull(workSheets[0].Rows[1].Values[4]);

            excelReader.Close();
        }

        //TESTS NOT USED AS DATATABLE NOT AVAILABLE IN SILVERLIGHT, MAY NEED TO USE BINDABLEDATAGRID PROJECT AND ADAPT EXCELDATAREADER
        //TO USE THIS?
        //[TestMethod]
        //[Tag("Xml")]
        //public void DataReader_Read_Test()
        //{
        //    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/Test_num_double_date_bool_string.xlsx"));

        //    DataTable table = new DataTable();
        //    table.Columns.Add(new DataColumn("num_col", typeof(int)));
        //    table.Columns.Add(new DataColumn("double_col", typeof(double)));
        //    table.Columns.Add(new DataColumn("date_col", typeof(DateTime)));
        //    table.Columns.Add(new DataColumn("boo_col", typeof(bool)));

        //    int fieldCount = -1;

        //    while (excelReader.Read())
        //    {
        //        fieldCount = excelReader.FieldCount;
        //        table.Rows.Add(excelReader.GetInt32(0), excelReader.GetDouble(1), excelReader.GetDateTime(2), excelReader.IsDBNull(4));
        //    }

        //    excelReader.Close();

        //    Assert.AreEqual(6, fieldCount);

        //    Assert.AreEqual(30, table.Rows.Count);

        //    Assert.AreEqual(1, int.Parse(table.Rows[0][0].ToString()));
        //    Assert.AreEqual(1346269, int.Parse(table.Rows[29][0].ToString()));

        //    //double + Formula
        //    Assert.AreEqual(1.02, double.Parse(table.Rows[0][1].ToString()));
        //    Assert.AreEqual(4.08, double.Parse(table.Rows[2][1].ToString()));
        //    Assert.AreEqual(547608330.24, double.Parse(table.Rows[29][1].ToString()));

        //    //Date + Formula
        //    Assert.AreEqual(new DateTime(2009, 5, 11).ToShortDateString(), DateTime.Parse(table.Rows[0][2].ToString()).ToShortDateString());
        //    Assert.AreEqual(new DateTime(2009, 11, 30).ToShortDateString(), DateTime.Parse(table.Rows[29][2].ToString()).ToShortDateString());
        //}


        //[TestMethod]
        //[Tag("Xml")]
        //public void DataReader_NextResult_Test()
        //{
        //    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("TestMultiSheet"));

        //    Assert.AreEqual(3, excelReader.ResultsCount);

        //    DataTable table = new DataTable();
        //    table.Columns.Add("c1", typeof(int)); table.Columns.Add("c2", typeof(int)); table.Columns.Add("c3", typeof(int)); table.Columns.Add("c4", typeof(int));

        //    int fieldCount = -1;

        //    while (excelReader.Read())
        //    {
        //        fieldCount = excelReader.FieldCount;
        //        table.Rows.Add(excelReader.GetInt32(0), excelReader.GetInt32(1), excelReader.GetInt32(2), excelReader.GetInt32(3));
        //    }

        //    Assert.AreEqual(12, table.Rows.Count);
        //    Assert.AreEqual(4, fieldCount);
        //    Assert.AreEqual(1, table.Rows[11][3]);

        //    excelReader.NextResult();
        //    table.Rows.Clear();

        //    while (excelReader.Read())
        //    {
        //        fieldCount = excelReader.FieldCount;
        //        table.Rows.Add(excelReader.GetInt32(0), excelReader.GetInt32(1), excelReader.GetInt32(2), excelReader.GetInt32(3));
        //    }

        //    Assert.AreEqual(12, table.Rows.Count);
        //    Assert.AreEqual(4, fieldCount);
        //    Assert.AreEqual(2, table.Rows[11][3]);


        //    excelReader.NextResult();
        //    table.Rows.Clear();

        //    while (excelReader.Read())
        //    {
        //        fieldCount = excelReader.FieldCount;
        //        table.Rows.Add(excelReader.GetInt32(0), excelReader.GetInt32(1));
        //    }

        //    Assert.AreEqual(5, table.Rows.Count);
        //    Assert.AreEqual(2, fieldCount);
        //    Assert.AreEqual(3, table.Rows[4][1]);

        //    Assert.AreEqual(false, excelReader.NextResult());

        //    excelReader.Close();
        //}

        [TestMethod]
        [Tag("Xml")]
        public void Test_Toyota()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/AMAXX.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            Assert.AreEqual("TOYOTA", workSheets[0].Rows[7].Values[5]);
        }

        [TestMethod]
        [Tag("Xml")]
        public void Test_Genes()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(GetTestWorkBook("Resources/Genes.xlsx"));

            excelReader.WorkBookFactory = new ExcelWorkBookFactory();

            IWorkBook excelWorkBook = excelReader.AsWorkBook();
            List<IWorkSheet> workSheets = (List<IWorkSheet>)excelWorkBook.WorkSheets;

            Assert.AreEqual(101526.9, double.Parse(workSheets[0].Rows[5].Values[6].ToString()));
            Assert.AreEqual(38713.79, double.Parse(workSheets[0].Rows[24].Values[7].ToString()));

            Assert.AreEqual("101526.9", workSheets[0].Rows[5].Values[6]);
            Assert.AreEqual("38713.79", workSheets[0].Rows[24].Values[7]);

        }
    }
}