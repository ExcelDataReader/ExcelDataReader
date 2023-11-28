using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

using NUnit.Framework;

namespace ExcelDataReader.Tests
{
    public abstract class ExcelTestBase
    {
        protected IExcelDataReader OpenReader(string name)
        {
            return OpenReader(OpenStream(name));
        }

        protected abstract Stream OpenStream(string name);

        protected abstract IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null);

        [Test]
        public void IssueDateAndTime1468Test()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Encoding_Formula_Date_1520");
            DataSet dataSet = excelReader.AsDataSet();

            string val1 = new DateTime(2009, 05, 01).ToShortDateString();
            string val2 = DateTime.Parse(dataSet.Tables[0].Rows[1][1].ToString()).ToShortDateString();

            Assert.AreEqual(val1, val2);

            val1 = new DateTime(2009, 1, 1, 11, 0, 0).ToShortTimeString();
            val2 = DateTime.Parse(dataSet.Tables[0].Rows[2][4].ToString()).ToShortTimeString();

            Assert.AreEqual(val1, val2);
        }

        [Test]
        public void Issue11773Exponential()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Issue_11773_Exponential");
            var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

            Assert.AreEqual(2566.3716814159293D, dataSet.Tables[0].Rows[0][6]);
        }

        [Test]
        public void Issue11773ExponentialCommas()
        {
#if NETCOREAPP1_0
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");
#else
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE", false);
#endif

            using IExcelDataReader excelReader = OpenReader("Test_Issue_11773_Exponential");
            var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

            Assert.AreEqual(2566.3716814159293D, dataSet.Tables[0].Rows[0][6]);
        }

        /// <summary>
        /// Makes sure that we can read data from the first row of last sheet
        /// </summary>
        [Test]
        public void Issue12271NextResultSet()
        {
            using IExcelDataReader excelReader = OpenReader("Test_LotsOfSheets");
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

        [Test]
        public void AsDataSetTestReadSheetNames()
        {
            using IExcelDataReader reader = OpenReader("TestOpen");
            Assert.AreEqual(3, reader.ResultsCount);

            DataSet dataSet = reader.AsDataSet();

            Assert.IsTrue(dataSet != null);
            Assert.AreEqual(3, dataSet.Tables.Count);
            Assert.AreEqual(7, dataSet.Tables["Sheet1"].Rows.Count);
            Assert.AreEqual(11, dataSet.Tables["Sheet1"].Columns.Count);
        }

        [Test]
        public void AsDataSetTest()
        {
            using IExcelDataReader excelReader = OpenReader("TestChess");
            DataSet result = excelReader.AsDataSet();

            Assert.IsTrue(result != null);
            Assert.AreEqual(1, result.Tables.Count);
            Assert.AreEqual(4, result.Tables[0].Rows.Count);
            Assert.AreEqual(6, result.Tables[0].Columns.Count);

            Assert.AreEqual(1, result.Tables[0].Rows[3][5]);
            Assert.AreEqual(1, result.Tables[0].Rows[2][0]);
        }

        [Test]
        public void AsDataSetTestRowCount()
        {
            using IExcelDataReader excelReader = OpenReader("TestChess");
            var result = excelReader.AsDataSet(Configuration.NoColumnNamesConfiguration);

            Assert.AreEqual(4, result.Tables[0].Rows.Count);
        }

        [Test]
        public void AsDataSetTestRowCountFirstRowAsColumnNames()
        {
            using IExcelDataReader excelReader = OpenReader("TestChess");
            var result = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

            Assert.AreEqual(3, result.Tables[0].Rows.Count);
        }

        [Test]
        public void ColumnWidthsTest()
        {
            // XLSX was manually edited to include a <col></col> element with closing tag
            using var reader = OpenReader("ColumnWidthsTest");
            reader.Read();

            // The expected values do not quite match what you see in Excel, is that correct?
            Assert.AreEqual(8.43, reader.GetColumnWidth(0));
            Assert.AreEqual(0, reader.GetColumnWidth(1));
            Assert.AreEqual(15.140625, reader.GetColumnWidth(2));
            Assert.AreEqual(28.7109375, reader.GetColumnWidth(3));

            var expectedException = typeof(ArgumentException);
            var exception = Assert.Throws(expectedException, () =>
            {
                reader.GetColumnWidth(4);
            });

#if NET5_0_OR_GREATER
                Assert.AreEqual($"Column at index 4 does not exist. (Parameter 'i')",
                    exception.Message);
#else
            Assert.AreEqual($"Column at index 4 does not exist.{Environment.NewLine}Parameter name: i",
                exception.Message);
#endif
        }

        [Test]
        public void GitIssue323DoubleClose()
        {
            using var reader = OpenReader("Test10x10");
            reader.Read();
            reader.Close();
        }

        [Test]
        public void MergedCells()
        {
            // XLSX was manually edited to include a <mergecell></mergecell> element with closing tag
            using var excelReader = OpenReader("Test_MergedCell");
            excelReader.Read();

            Assert.That(excelReader.MergeCells, Is.EquivalentTo(new[] {
                    new[] { 1, 2, 0, 1 },
                    new[] { 1, 5, 2, 2 },
                    new[] { 3, 5, 0, 0 },
                    new[] { 6, 6, 0, 2 },
                }).Using<CellRange, int[]>((a, e) => a.FromRow == e[0] && a.ToRow == e[1] && a.FromColumn == e[2] && a.ToColumn == e[3]));
        }

        [Test]
        public void OpenXmlLeaveOpen()
        {
            // Verify the file stream is closed and disposed by the reader
            {
                var stream = OpenStream("Test10x10");
                using (IExcelDataReader excelReader = OpenReader(stream, new ExcelReaderConfiguration()
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

        [Test]
        public void RowHeight()
        {
            using var reader = OpenReader("CollapsedHide");
            // Verify the row heights are set when expected, and converted to points from twips
            reader.Read();
            Assert.AreEqual(15, reader.RowHeight);

            reader.Read();
            Assert.AreEqual(38.25, reader.RowHeight);

            reader.Read();
            Assert.AreEqual(6, reader.RowHeight);

            reader.Read();
            Assert.AreEqual(0, reader.RowHeight);
        }

        [Test]
        public void GitIssue245NoCodeName()
        {
            // Test no CodeName = null
            using var reader = OpenReader("Test10x10");
            Assert.AreEqual(null, reader.CodeName);
        }

        [Test]
        public void GitIssue245CodeName()
        {
            // Test CodeName is set
            using var reader = OpenReader("Test_Excel_Dataset");
            Assert.AreEqual("Sheet1", reader.CodeName);
        }

        [Test]
        public void GitIssue241Simple()
        {
            using var reader = OpenReader("Test_git_issue_224_simple");
            Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft åäö &T&CCenter åäö &D&RRight  åäö &P"), "Header");
            Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft åäö &P&CFooter åäö &P&RRight åäö &D"), "Footer");
        }

        [Test]
        public void Dimension10X10000Test()
        {
            using IExcelDataReader excelReader = OpenReader("Test10x10000");
            DataTable result = excelReader.AsDataSet().Tables[0];

            Assert.AreEqual(10000, result.Rows.Count);
            Assert.AreEqual(10, result.Columns.Count);
            Assert.AreEqual("1x2", result.Rows[1][1]);
            Assert.AreEqual("1x10", result.Rows[1][9]);
            Assert.AreEqual("1x1", result.Rows[9999][0]);
            Assert.AreEqual("1x10", result.Rows[9999][9]);
        }


        [Test]
        public void Dimension10X10Test()
        {
            using IExcelDataReader excelReader = OpenReader("Test10x10");
            DataTable result = excelReader.AsDataSet().Tables[0];

            Assert.AreEqual(10, result.Rows.Count);
            Assert.AreEqual(10, result.Columns.Count);
            Assert.AreEqual("10x10", result.Rows[1][0]);
            Assert.AreEqual("10x27", result.Rows[9][9]);
        }

        [Test]
        public void Dimension255X10Test()
        {
            using IExcelDataReader excelReader = OpenReader("Test255x10");
            DataTable result = excelReader.AsDataSet().Tables[0];

            Assert.AreEqual(10, result.Rows.Count);
            Assert.AreEqual(255, result.Columns.Count);
            Assert.AreEqual("1", result.Rows[9][254].ToString());
            Assert.AreEqual("one", result.Rows[1][1].ToString());
        }

        [Test]
        public void DoublePrecisionTest()
        {
            using IExcelDataReader excelReader = OpenReader("TestDoublePrecision");
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

        protected abstract DateTime GitIssue82TodayDate { get; }

        [Test]
        public void GitIssue82Date1900()
        {
            using var excelReader = OpenReader("roo_1900_base");
            // 15/06/2009
            // 4/19/2013 (=TODAY() when file was saved)

            DataSet result = excelReader.AsDataSet();
            Assert.AreEqual(new DateTime(2009, 6, 15), (DateTime)result.Tables[0].Rows[0][0]);
            Assert.AreEqual(GitIssue82TodayDate, (DateTime)result.Tables[0].Rows[1][0]);
        }

        [Test]
        public void GitIssue82Date1904()
        {
            using var excelReader = OpenReader("roo_1904_base");
            // 15/06/2009
            // 4/19/2013 (=TODAY() when file was saved)

            DataSet result = excelReader.AsDataSet();
            Assert.AreEqual(new DateTime(2009, 6, 15), (DateTime)result.Tables[0].Rows[0][0]);
            Assert.AreEqual(GitIssue82TodayDate, (DateTime)result.Tables[0].Rows[1][0]);
        }

        [Test]
        public void TestBlankHeader()
        {
            using IExcelDataReader excelReader = OpenReader("Test_BlankHeader");
            excelReader.Read();
            Assert.AreEqual(4, excelReader.FieldCount);
            excelReader.Read();
        }

        [Test]
        public void IssueDecimal1109Test()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Decimal_1109");
            DataSet dataSet = excelReader.AsDataSet();

            Assert.AreEqual(3.14159, dataSet.Tables[0].Rows[0][0]);

            const double val1 = -7080.61;
            double val2 = (double)dataSet.Tables[0].Rows[0][1];
            Assert.AreEqual(val1, val2);
        }
        
        [Test]
        public void IssueEncoding1520Test()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Encoding_Formula_Date_1520");
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

        [Test]
        public void TestIssue11601ReadSheetNames()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Excel_Dataset");
            Assert.AreEqual("test.csv", excelReader.Name);

            excelReader.NextResult();
            Assert.AreEqual("Sheet2", excelReader.Name);

            excelReader.NextResult();
            Assert.AreEqual("Sheet3", excelReader.Name);
        }

        [Test]
        public void GitIssue250RichText()
        {
            using var reader = OpenReader("Test_git_issue_250_richtext");
            reader.Read();
            var text = reader.GetString(0);
            Assert.AreEqual("Lorem ipsum dolor sit amet, ei pri verterem efficiantur, per id meis idque deterruisset.", text);
        }

        [Test]
        public void GitIssue270EmptyRowsAtTheEnd()
        {
            // AsDataSet() trims trailing blank rows
            using (var reader = OpenReader("Test_git_issue_270"))
            {
                var dataSet = reader.AsDataSet();
                Assert.AreEqual(1, dataSet.Tables[0].Rows.Count);
            }

            // Reader methods do not trim trailing blank rows
            using (var reader = OpenReader("Test_git_issue_270"))
            {
                var rowCount = 0;
                while (reader.Read())
                    rowCount++;
                Assert.AreEqual(65536, rowCount);
            }
        }
        
        [Test]
        public void GitIssue283TimeSpan()
        {
            using var reader = OpenReader("Test_git_issue_283_TimeSpan");
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

        [Test]
        public void GitIssue329Error()
        {
            using var reader = OpenReader("Test_git_issue_329_error");
            var result = reader.AsDataSet().Tables[0];

            // AsDataSet trims trailing empty rows
            Assert.AreEqual(0, result.Rows.Count);

            // Check errors on first row return null
            reader.Read();
            Assert.IsNull(reader.GetValue(0));
            Assert.AreEqual(CellError.DIV0, reader.GetCellError(0));

            Assert.IsNull(reader.GetValue(1));
            Assert.AreEqual(CellError.NA, reader.GetCellError(1));

            Assert.IsNull(reader.GetValue(2));
            Assert.AreEqual(CellError.VALUE, reader.GetCellError(2));

            Assert.AreEqual(1, reader.RowCount);
        }

        [Test]
        public void Issue4031NullColumn()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Issue_4031_NullColumn");
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

        [Test]
        public void Issue7433IllegalOleAutDate()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Issue_7433_IllegalOleAutDate");
            DataSet dataSet = excelReader.AsDataSet();

            Assert.AreEqual(3.10101195608231E+17, dataSet.Tables[0].Rows[0][0]);
            Assert.AreEqual("B221055625", dataSet.Tables[0].Rows[1][0]);
            Assert.AreEqual(4.12721197309241E+17, dataSet.Tables[0].Rows[2][0]);
        }

        [Test]
        public void Issue8536Test()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Issue_8536");
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

        [Test]
        public void Issue10725()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Issue_10725");
            excelReader.Read();
            Assert.AreEqual(8.8, excelReader.GetValue(0));

            DataSet result = excelReader.AsDataSet();

            Assert.AreEqual(8.8, result.Tables[0].Rows[0][0]);
        }

        [Test]
        public void Issue11397CurrencyTest()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Issue_11397");
            DataSet dataSet = excelReader.AsDataSet();

            Assert.AreEqual("$44.99", dataSet.Tables[0].Rows[1][0].ToString()); // general in spreadsheet so should be a string including the $
            Assert.AreEqual(44.99, double.Parse(dataSet.Tables[0].Rows[2][0].ToString())); // currency euros in spreadsheet so should be a currency
            Assert.AreEqual(44.99, double.Parse(dataSet.Tables[0].Rows[3][0].ToString())); // currency pounds in spreadsheet so should be a currency
        }

        [Test]
        public void Issue11435Colors()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Issue_11435_Colors");
            DataSet dataSet = excelReader.AsDataSet();

            Assert.AreEqual("test1", dataSet.Tables[0].Rows[0][0].ToString());
            Assert.AreEqual("test2", dataSet.Tables[0].Rows[0][1].ToString());
            Assert.AreEqual("test3", dataSet.Tables[0].Rows[0][2].ToString());

            excelReader.Read();

            Assert.AreEqual("test1", excelReader.GetString(0));
            Assert.AreEqual("test2", excelReader.GetString(1));
            Assert.AreEqual("test3", excelReader.GetString(2));
        }

        [Test]
        public void Issue11479BlankSheet()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Issue_11479_BlankSheet");
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

        [Test]
        public void Issue11573BlankValues()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Issue_11573_BlankValues");
            var dataSet = excelReader.AsDataSet();

            Assert.AreEqual(1D, dataSet.Tables[0].Rows[12][0]);
            Assert.AreEqual("070202", dataSet.Tables[0].Rows[12][1]);
        }

        [Test]
        public void IssueBoolFormula()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Issue_BoolFormula");
            DataSet dataSet = excelReader.AsDataSet();

            Assert.AreEqual(true, dataSet.Tables[0].Rows[0][0]);
        }

        [Test]
        public void IssueDateFormatButNotDate()
        {
            // we want to make sure that if a cell is formatted as a date but it's contents are not a date then
            // the output is not a date (it was ending up as datetime.min)
            using IExcelDataReader excelReader = OpenReader("Test_Issue_DateFormatButNotDate");
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

        [Test]
        public void DataReaderReadTest()
        {
            using IExcelDataReader r = OpenReader("Test_num_double_date_bool_string");
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

        [Test]
        public void MultiSheetTest()
        {
            using IExcelDataReader excelReader = OpenReader("TestMultiSheet");
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

        [Test]
        public void DataReaderNextResultTest()
        {
            using IExcelDataReader r = OpenReader("TestMultiSheet");
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

        [Test]
        public void UnicodeCharsTest()
        {
            using IExcelDataReader excelReader = OpenReader("TestUnicodeChars");
            DataTable result = excelReader.AsDataSet().Tables[0];

            Assert.AreEqual(3, result.Rows.Count);
            Assert.AreEqual(8, result.Columns.Count);
            Assert.AreEqual("\u00e9\u0417", result.Rows[1][0].ToString());
        }

        [Test]
        public void GitIssue29ReadSheetStatesReadsCorrectly()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Excel_Dataset");
            Assert.AreEqual("hidden", excelReader.VisibleState);

            excelReader.NextResult();
            Assert.AreEqual("visible", excelReader.VisibleState);

            excelReader.NextResult();
            Assert.AreEqual("veryhidden", excelReader.VisibleState);
        }

        [Test]
        public void GitIssue29AsDataSetProvidesCorrectSheetState()
        {
            using IExcelDataReader reader = OpenReader("Test_Excel_Dataset");
            var dataSet = reader.AsDataSet();

            Assert.IsTrue(dataSet != null);
            Assert.AreEqual(3, dataSet.Tables.Count);
            Assert.AreEqual("hidden", dataSet.Tables[0].ExtendedProperties["visiblestate"]);
            Assert.AreEqual("visible", dataSet.Tables[1].ExtendedProperties["visiblestate"]);
            Assert.AreEqual("veryhidden", dataSet.Tables[2].ExtendedProperties["visiblestate"]);
        }

        [Test]
        public void GitIssue389FilterSheetByVisibility()
        {
            using IExcelDataReader excelReader = OpenReader("Test_Excel_Dataset");
            var result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                FilterSheet = (r, index) => r.VisibleState == "visible"
            });

            Assert.AreEqual(1, result.Tables.Count);
        }

        [Test]
        public void TestNumDoubleDateBoolString()
        {
            using IExcelDataReader excelReader = OpenReader("Test_num_double_date_bool_string");
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

        [Test]
        public void GitIssue160FilterRow()
        {
            // Check there are four rows with data, including empty and hidden rows
            using (var reader = OpenReader("CollapsedHide"))
            {
                var dataSet = reader.AsDataSet();

                Assert.AreEqual(4, dataSet.Tables[0].Rows.Count);
            }

            // Check there are two rows with content
            using (var reader = OpenReader("CollapsedHide"))
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
            using (var reader = OpenReader("CollapsedHide"))
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
        }

        [Test]
        public void GitIssue300FilterColumn()
        {
            // Check there are two columns with data
            using (var reader = OpenReader("CollapsedHide"))
            {
                var dataSet = reader.AsDataSet();

                Assert.AreEqual(2, dataSet.Tables[0].Columns.Count);
            }

            // Check there is one column when skipping the first
            using (var reader = OpenReader("CollapsedHide"))
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

        [Test]
        public void GitIssue483CellErrorEmptyRow()
        {
            // Check there are four rows with no errors and no NREs
            using var reader = OpenReader("CollapsedHide");
            reader.Read();
            Assert.AreEqual(null, reader.GetCellError(0));
            Assert.AreEqual(null, reader.GetCellError(1));

            reader.Read();
            Assert.AreEqual(null, reader.GetCellError(0));
            Assert.AreEqual(null, reader.GetCellError(1));

            reader.Read();
            Assert.AreEqual(null, reader.GetCellError(0));
            Assert.AreEqual(null, reader.GetCellError(1));

            reader.Read();
            Assert.AreEqual(null, reader.GetCellError(0));
            Assert.AreEqual(null, reader.GetCellError(1));
        }

        [Test]
        public void GitIssue532TrimEmptyColumns()
        {
            using var reader = OpenReader("Test_git_issue_532");
            while (reader.Read())
            {
                Assert.AreEqual(3, reader.FieldCount);
            }
        }
    }
}
