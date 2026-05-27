using System.Data;
using System.Globalization;

namespace ExcelDataReader.Tests;

public abstract class ExcelTestBase
{
    protected abstract DateTime Issue82_TodayDate { get; }

    [Test]
    public void IssueDateAndTime1468Test()
    {
        using IExcelDataReader excelReader = OpenReader("EncodingFormulaDate1520");
        DataSet dataSet = excelReader.AsDataSet();

        string val1 = new DateTime(2009, 05, 01).ToShortDateString();
        string val2 = DateTime.Parse(dataSet.Tables[0].Rows[1][1].ToString()).ToShortDateString();

        Assert.That(val2, Is.EqualTo(val1));

        val1 = new DateTime(2009, 1, 1, 11, 0, 0).ToShortTimeString();
        val2 = DateTime.Parse(dataSet.Tables[0].Rows[2][4].ToString()).ToShortTimeString();

        Assert.That(val2, Is.EqualTo(val1));
    }

    [Test]
    public void Issue11773_Exponential()
    {
        using IExcelDataReader excelReader = OpenReader("OldIssue11773_Exponential");
        var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

        Assert.That(dataSet.Tables[0].Rows[0][6], Is.EqualTo(2566.3716814159293D));
    }

    [Test]
    public void Issue11773_ExponentialCommas()
    {
#if NETCOREAPP1_0
        CultureInfo.CurrentCulture = new CultureInfo("de-DE");
#else
        System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE", false);
#endif

        using IExcelDataReader excelReader = OpenReader("OldIssue11773_Exponential");
        var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

        Assert.That(dataSet.Tables[0].Rows[0][6], Is.EqualTo(2566.3716814159293D));
    }

    /// <summary>
    /// Makes sure that we can read data from the first row of last sheet.
    /// </summary>
    [Test]
    public void Issue12271_NextResultSet()
    {
        using IExcelDataReader excelReader = OpenReader("LotsOfSheets");
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

            Assert.That(excelReader.GetString(0), Is.EqualTo("StaffName"));
        }
        while (excelReader.NextResult());
    }

    [Test]
    public void AsDataSetTestReadSheetNames()
    {
        using IExcelDataReader reader = OpenReader("Open");
        Assert.That(reader.ResultsCount, Is.EqualTo(3));

        DataSet dataSet = reader.AsDataSet();

        Assert.That(dataSet != null, Is.True);
        Assert.That(dataSet.Tables.Count, Is.EqualTo(3));
        Assert.That(dataSet.Tables["Sheet1"].Rows.Count, Is.EqualTo(7));
        Assert.That(dataSet.Tables["Sheet1"].Columns.Count, Is.EqualTo(11));
    }

    [Test]
    public void AsDataSetTest()
    {
        using IExcelDataReader excelReader = OpenReader("Chess");
        DataSet result = excelReader.AsDataSet();

        Assert.That(result != null, Is.True);
        Assert.That(result.Tables.Count, Is.EqualTo(1));
        Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(4));
        Assert.That(result.Tables[0].Columns.Count, Is.EqualTo(6));

        Assert.That(result.Tables[0].Rows[3][5], Is.EqualTo(1));
        Assert.That(result.Tables[0].Rows[2][0], Is.EqualTo(1));
    }

    [Test]
    public void AsDataSetTestRowCount()
    {
        using IExcelDataReader excelReader = OpenReader("Chess");
        var result = excelReader.AsDataSet(Configuration.NoColumnNamesConfiguration);

        Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(4));
    }

    [Test]
    public void AsDataSetTestRowCountFirstRowAsColumnNames()
    {
        using IExcelDataReader excelReader = OpenReader("Chess");
        var result = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

        Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(3));
    }

    [Test]
    public void ColumnWidthsTest()
    {
        // XLSX was manually edited to include a <col></col> element with closing tag
        using var reader = OpenReader("ColumnWidthsTest");
        reader.Read();

        // The expected values do not quite match what you see in Excel, is that correct?
        Assert.That(reader.GetColumnWidth(0), Is.EqualTo(8.43));
        Assert.That(reader.GetColumnWidth(1), Is.EqualTo(0));
        Assert.That(reader.GetColumnWidth(2), Is.EqualTo(15.140625));
        Assert.That(reader.GetColumnWidth(3), Is.EqualTo(28.7109375));

        var expectedException = typeof(ArgumentException);
        var exception = Assert.Throws(expectedException, () =>
        {
            reader.GetColumnWidth(4);
        });

        Assert.That(exception.Message, Does.StartWith($"Column at index 4 does not exist."));
    }

    [Test]
    public void Issue323_DoubleClose()
    {
        using var reader = OpenReader("10x10");
        reader.Read();
        reader.Close();
    }

    [Test]
    public void MergedCells()
    {
        // XLSX was manually edited to include a <mergecell></mergecell> element with closing tag
        using var excelReader = OpenReader("MergedCell");
        excelReader.Read();

        Assert.That(excelReader.MergeCells, Is.EquivalentTo(new[]
        {
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
            var stream = OpenStream("10x10");
            using (IExcelDataReader excelReader = OpenReader(stream, new ExcelReaderConfiguration()
            {
                LeaveOpen = true
            }))
            {
                _ = excelReader.AsDataSet();
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
        Assert.That(reader.RowHeight, Is.EqualTo(15));

        reader.Read();
        Assert.That(reader.RowHeight, Is.EqualTo(38.25));

        reader.Read();
        Assert.That(reader.RowHeight, Is.EqualTo(6));

        reader.Read();
        Assert.That(reader.RowHeight, Is.EqualTo(0));
    }

    [Test]
    public void Issue245_NoCodeName()
    {
        // Test no CodeName = null
        using var reader = OpenReader("10x10");
        Assert.That(reader.CodeName, Is.EqualTo(null));
    }

    [Test]
    public void Issue245_CodeName()
    {
        // Test CodeName is set
        using var reader = OpenReader("ExcelDataset");
        Assert.That(reader.CodeName, Is.EqualTo("Sheet1"));
    }

    [Test]
    public void Issue241_Simple()
    {
        using var reader = OpenReader("Issue224_Simple");
        Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft åäö &T&CCenter åäö &D&RRight  åäö &P"), "Header");
        Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft åäö &P&CFooter åäö &P&RRight åäö &D"), "Footer");
    }

    [Test]
    public void Dimension10X10000Test()
    {
        using IExcelDataReader excelReader = OpenReader("10x10000");
        DataTable result = excelReader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(10000));
        Assert.That(result.Columns.Count, Is.EqualTo(10));
        Assert.That(result.Rows[1][1], Is.EqualTo("1x2"));
        Assert.That(result.Rows[1][9], Is.EqualTo("1x10"));
        Assert.That(result.Rows[9999][0], Is.EqualTo("1x1"));
        Assert.That(result.Rows[9999][9], Is.EqualTo("1x10"));
    }

    [Test]
    public void Dimension10X10Test()
    {
        using IExcelDataReader excelReader = OpenReader("10x10");
        DataTable result = excelReader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(10));
        Assert.That(result.Columns.Count, Is.EqualTo(10));
        Assert.That(result.Rows[1][0], Is.EqualTo("10x10"));
        Assert.That(result.Rows[9][9], Is.EqualTo("10x27"));
    }

    [Test]
    public void Dimension255X10Test()
    {
        using IExcelDataReader excelReader = OpenReader("255x10");
        DataTable result = excelReader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(10));
        Assert.That(result.Columns.Count, Is.EqualTo(255));
        Assert.That(result.Rows[9][254].ToString(), Is.EqualTo("1"));
        Assert.That(result.Rows[1][1].ToString(), Is.EqualTo("one"));
    }

    [Test]
    public void DoublePrecisionTest()
    {
        using IExcelDataReader excelReader = OpenReader("DoublePrecision");
        DataTable result = excelReader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(10));

        const double excelPi = 3.14159265358979;

        Assert.That(result.Rows[2][1], Is.EqualTo(+excelPi));
        Assert.That(result.Rows[3][1], Is.EqualTo(-excelPi));

        Assert.That(result.Rows[4][1], Is.EqualTo(+excelPi * 1.0e-300));
        Assert.That(result.Rows[5][1], Is.EqualTo(-excelPi * 1.0e-300));

        Assert.That((double)result.Rows[6][1], Is.EqualTo(+excelPi * 1.0e300).Within(1e286)); // only accurate to 1e286 because excel only has 15 digits precision
        Assert.That((double)result.Rows[7][1], Is.EqualTo(-excelPi * 1.0e300).Within(1e286));

        Assert.That(result.Rows[8][1], Is.EqualTo(+excelPi * 1.0e14));
        Assert.That(result.Rows[9][1], Is.EqualTo(-excelPi * 1.0e14));
    }

    [Test]
    public void Issue82_Date1900()
    {
        // 15/06/2009
        // 4/19/2013 (=TODAY() when file was saved)
        using var excelReader = OpenReader("roo_1900_base");

        DataSet result = excelReader.AsDataSet();
        Assert.That((DateTime)result.Tables[0].Rows[0][0], Is.EqualTo(new DateTime(2009, 6, 15)));
        Assert.That((DateTime)result.Tables[0].Rows[1][0], Is.EqualTo(Issue82_TodayDate));
    }

    [Test]
    public void Issue82_Date1904()
    {
        // 15/06/2009
        // 4/19/2013 (=TODAY() when file was saved)
        using var excelReader = OpenReader("roo_1904_base");

        DataSet result = excelReader.AsDataSet();
        Assert.That((DateTime)result.Tables[0].Rows[0][0], Is.EqualTo(new DateTime(2009, 6, 15)));
        Assert.That((DateTime)result.Tables[0].Rows[1][0], Is.EqualTo(Issue82_TodayDate));
    }

    [Test]
    public void TestBlankHeader()
    {
        using IExcelDataReader excelReader = OpenReader("BlankHeader");
        excelReader.Read();
        Assert.That(excelReader.FieldCount, Is.EqualTo(4));
        excelReader.Read();
    }

    [Test]
    public void IssueDecimal1109Test()
    {
        using IExcelDataReader excelReader = OpenReader("Decimal1109");
        DataSet dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo(3.14159));

        const double val1 = -7080.61;
        double val2 = (double)dataSet.Tables[0].Rows[0][1];
        Assert.That(val2, Is.EqualTo(val1));
    }

    [Test]
    public void IssueEncoding1520Test()
    {
        using IExcelDataReader excelReader = OpenReader("EncodingFormulaDate1520");
        DataSet dataSet = excelReader.AsDataSet();

        string val1 = "Simon Hodgetts";
        string val2 = dataSet.Tables[0].Rows[2][0].ToString();
        Assert.That(val2, Is.EqualTo(val1));

        val1 = "John test";
        val2 = dataSet.Tables[0].Rows[1][0].ToString();
        Assert.That(val2, Is.EqualTo(val1));

        // librement réutilisable
        val1 = "librement réutilisable";
        val2 = dataSet.Tables[0].Rows[7][0].ToString();
        Assert.That(val2, Is.EqualTo(val1));

        val2 = dataSet.Tables[0].Rows[8][0].ToString();
        Assert.That(val2, Is.EqualTo(val1));
    }

    [Test]
    public void TestIssue11601_ReadSheetNames()
    {
        using IExcelDataReader excelReader = OpenReader("ExcelDataset");
        Assert.That(excelReader.Name, Is.EqualTo("test.csv"));

        excelReader.NextResult();
        Assert.That(excelReader.Name, Is.EqualTo("Sheet2"));

        excelReader.NextResult();
        Assert.That(excelReader.Name, Is.EqualTo("Sheet3"));
    }

    [Test]
    public void Issue250_RichText()
    {
        using var reader = OpenReader("Issue250_Richtext");
        reader.Read();
        var text = reader.GetString(0);
        Assert.That(text, Is.EqualTo("Lorem ipsum dolor sit amet, ei pri verterem efficiantur, per id meis idque deterruisset."));
    }

    [Test]
    public void Issue270_EmptyRowsAtTheEnd()
    {
        // AsDataSet() trims trailing blank rows
        using (var reader = OpenReader("Issue270"))
        {
            var dataSet = reader.AsDataSet();
            Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(1));
        }

        // Reader methods do not trim trailing blank rows
        using (var reader = OpenReader("Issue270"))
        {
            var rowCount = 0;
            while (reader.Read())
                rowCount++;
            Assert.That(rowCount, Is.EqualTo(65536));
        }
    }

    [Test]
    public void Issue283_IsoFormatTimeSpan()
    {
        using var reader = OpenReader("Issue283_TimeSpan");
        reader.Read();
        Assert.That(new TimeSpan(0), Is.EqualTo((TimeSpan)reader[0]));
        Assert.That(new DateTime(1899, 12, 31), Is.EqualTo((DateTime)reader[2])); // Excel says 1/0/1900, not valid in .NET

        reader.Read();
        Assert.That(new TimeSpan(1, 0, 0, 0, 0), Is.EqualTo((TimeSpan)reader[0]));
        Assert.That(new DateTime(1900, 1, 1), Is.EqualTo((DateTime)reader[2]));

        reader.Read();
        Assert.That(new TimeSpan(2, 0, 0, 0, 0), Is.EqualTo((TimeSpan)reader[0]));

        reader.Read();
        Assert.That(new TimeSpan(0, 1392, 0, 0, 0), Is.EqualTo((TimeSpan)reader[0]));

        reader.Read();
        Assert.That(new TimeSpan(0, 1416, 0, 0, 0), Is.EqualTo((TimeSpan)reader[0]));
        Assert.That(new DateTime(1900, 2, 28), Is.EqualTo((DateTime)reader[2]));

        reader.Read();
        Assert.That(new TimeSpan(0, 1440, 0, 0, 0), Is.EqualTo((TimeSpan)reader[0]));
        Assert.That(new DateTime(1900, 2, 28), Is.EqualTo((DateTime)reader[2])); // Excel says 2/29/1900, not valid in .NET

        reader.Read();
        Assert.That(new TimeSpan(0, 1464, 0, 0, 0), Is.EqualTo((TimeSpan)reader[0]));
        Assert.That(new DateTime(1900, 3, 1), Is.EqualTo((DateTime)reader[2]));

        reader.Read();
        Assert.That(new TimeSpan(0, 1488, 0, 0, 0), Is.EqualTo((TimeSpan)reader[0]));

        reader.Read();
        Assert.That(new TimeSpan(0, 1512, 0, 0, 0), Is.EqualTo((TimeSpan)reader[0]));
    }

    [Test]
    public void Issue329_Error()
    {
        using var reader = OpenReader("Issue329_Error");
        var result = reader.AsDataSet().Tables[0];

        // AsDataSet trims trailing empty rows
        Assert.That(result.Rows.Count, Is.EqualTo(0));

        // Check errors on first row return null
        reader.Read();
        Assert.That(reader.GetValue(0), Is.Null);
        Assert.That(reader.GetCellError(0), Is.EqualTo(CellError.DIV0));

        Assert.That(reader.GetValue(1), Is.Null);
        Assert.That(reader.GetCellError(1), Is.EqualTo(CellError.NA));

        Assert.That(reader.GetValue(2), Is.Null);
        Assert.That(reader.GetCellError(2), Is.EqualTo(CellError.VALUE));

        Assert.That(reader.RowCount, Is.EqualTo(1));
    }

    [Test]
    public void Issue4031_NullColumn()
    {
        using IExcelDataReader excelReader = OpenReader("OldIssue4031_NullColumn");

        // DataSet dataSet = excelReader.AsDataSet(true);
        excelReader.Read();
        Assert.That(excelReader.GetValue(0), Is.Null);
        Assert.That(excelReader.GetString(1), Is.EqualTo("a"));
        Assert.That(excelReader.GetString(2), Is.EqualTo("b"));
        Assert.That(excelReader.GetValue(3), Is.Null);
        Assert.That(excelReader.GetString(4), Is.EqualTo("d"));

        excelReader.Read();
        Assert.That(excelReader.GetValue(0), Is.Null);
        Assert.That(excelReader.GetValue(1), Is.Null);
        Assert.That(excelReader.GetString(2), Is.EqualTo("Test"));
        Assert.That(excelReader.GetValue(3), Is.Null);
        Assert.That(excelReader.GetDouble(4), Is.EqualTo(1));
    }

    [Test]
    public void Issue7433_IllegalOleAutDate()
    {
        using IExcelDataReader excelReader = OpenReader("OldIssue7433_IllegalOleAutDate");
        DataSet dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo(3.10101195608231E+17));
        Assert.That(dataSet.Tables[0].Rows[1][0], Is.EqualTo("B221055625"));
        Assert.That(dataSet.Tables[0].Rows[2][0], Is.EqualTo(4.12721197309241E+17));
    }

    [Test]
    public void Issue8536_Test()
    {
        using IExcelDataReader excelReader = OpenReader("OldIssue8536");
        DataSet dataSet = excelReader.AsDataSet();

        // date
        var date1900 = dataSet.Tables[0].Rows[7][1];
        Assert.That(date1900.GetType(), Is.EqualTo(typeof(DateTime)));
        Assert.That(date1900, Is.EqualTo(new DateTime(1900, 1, 1)));

        // xml encoded chars
        var xmlChar1 = dataSet.Tables[0].Rows[6][1];
        Assert.That(xmlChar1.GetType(), Is.EqualTo(typeof(string)));
        Assert.That(xmlChar1, Is.EqualTo("&#x26; "));

        // number but matches a date serial
        var dateNum1 = dataSet.Tables[0].Rows[5][1];
        Assert.That(dateNum1.GetType(), Is.EqualTo(typeof(double)));
        Assert.That(double.Parse(dateNum1.ToString()), Is.EqualTo(41244));

        // date
        var date1 = dataSet.Tables[0].Rows[4][1];
        Assert.That(date1.GetType(), Is.EqualTo(typeof(DateTime)));
        Assert.That(date1, Is.EqualTo(new DateTime(2012, 12, 1)));

        // double
        var num1 = dataSet.Tables[0].Rows[3][1];
        Assert.That(num1.GetType(), Is.EqualTo(typeof(double)));
        Assert.That(double.Parse(num1.ToString()), Is.EqualTo(12345));

        // boolean issue
        var val2 = dataSet.Tables[0].Rows[2][1];
        Assert.That(val2.GetType(), Is.EqualTo(typeof(bool)));
        Assert.That((bool)val2, Is.True);
    }

    [Test]
    public void Issue10725()
    {
        using IExcelDataReader excelReader = OpenReader("OldIssue10725");
        excelReader.Read();
        Assert.That(excelReader.GetValue(0), Is.EqualTo(8.8));

        DataSet result = excelReader.AsDataSet();

        Assert.That(result.Tables[0].Rows[0][0], Is.EqualTo(8.8));
    }

    [Test]
    public void Issue11397_CurrencyTest()
    {
        using IExcelDataReader excelReader = OpenReader("OldIssue11397");
        DataSet dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[1][0].ToString(), Is.EqualTo("$44.99")); // general in spreadsheet so should be a string including the $
        Assert.That(double.Parse(dataSet.Tables[0].Rows[2][0].ToString()), Is.EqualTo(44.99)); // currency euros in spreadsheet so should be a currency
        Assert.That(double.Parse(dataSet.Tables[0].Rows[3][0].ToString()), Is.EqualTo(44.99)); // currency pounds in spreadsheet so should be a currency
    }

    [Test]
    public void Issue11435_Colors()
    {
        using IExcelDataReader excelReader = OpenReader("OldIssue11435_Colors");
        DataSet dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[0][0].ToString(), Is.EqualTo("test1"));
        Assert.That(dataSet.Tables[0].Rows[0][1].ToString(), Is.EqualTo("test2"));
        Assert.That(dataSet.Tables[0].Rows[0][2].ToString(), Is.EqualTo("test3"));

        excelReader.Read();

        Assert.That(excelReader.GetString(0), Is.EqualTo("test1"));
        Assert.That(excelReader.GetString(1), Is.EqualTo("test2"));
        Assert.That(excelReader.GetString(2), Is.EqualTo("test3"));
    }

    [Test]
    public void Issue11479_BlankSheet()
    {
        using IExcelDataReader excelReader = OpenReader("OldIssue11479_BlankSheet");

        // DataSet result = excelReader.AsDataSet(true);
        excelReader.Read();
        Assert.That(excelReader.FieldCount, Is.EqualTo(5));
        excelReader.NextResult();
        excelReader.Read();
        Assert.That(excelReader.FieldCount, Is.EqualTo(0));

        excelReader.NextResult();
        excelReader.Read();
        Assert.That(excelReader.FieldCount, Is.EqualTo(0));
    }

    [Test]
    public void Issue11573_BlankValues()
    {
        using IExcelDataReader excelReader = OpenReader("OldIssue11573_BlankValues");
        var dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[12][0], Is.EqualTo(1D));
        Assert.That(dataSet.Tables[0].Rows[12][1], Is.EqualTo("070202"));
    }

    [Test]
    public void IssueBoolFormula()
    {
        using IExcelDataReader excelReader = OpenReader("BoolFormula");
        DataSet dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo(true));
    }

    [Test]
    public void IssueDateFormatButNotDate()
    {
        // we want to make sure that if a cell is formatted as a date but it's contents are not a date then
        // the output is not a date (it was ending up as datetime.min)
        using IExcelDataReader excelReader = OpenReader("DateFormatButNotDate");
        excelReader.Read();
        Assert.That(excelReader.GetValue(0), Is.EqualTo("columna"));
        Assert.That(excelReader.GetValue(1), Is.EqualTo("columnb"));
        Assert.That(excelReader.GetValue(2), Is.EqualTo("columnc"));
        Assert.That(excelReader.GetValue(3), Is.EqualTo("columnd"));
        Assert.That(excelReader.GetValue(4), Is.EqualTo("columne"));

        excelReader.Read();
        Assert.That(excelReader.GetValue(0), Is.EqualTo(1D));
        Assert.That(excelReader.GetValue(1), Is.EqualTo(2D));
        Assert.That(excelReader.GetValue(2), Is.EqualTo(3D));
        var value = excelReader.GetValue(3);
        Assert.That(value, Is.EqualTo(new DateTime(2013, 12, 10)));
        Assert.That(excelReader.GetValue(4), Is.EqualTo("red"));

        excelReader.Read();
        Assert.That(excelReader.GetValue(4), Is.EqualTo("yellow"));
    }

    [Test]
    public void DataReaderReadTest()
    {
        using IExcelDataReader r = OpenReader("NumDoubleDateBoolString");
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

        Assert.That(fieldCount, Is.EqualTo(6));

        Assert.That(table.Rows.Count, Is.EqualTo(30));

        Assert.That(int.Parse(table.Rows[0][0].ToString()), Is.EqualTo(1));
        Assert.That(int.Parse(table.Rows[29][0].ToString()), Is.EqualTo(1346269));

        // double + Formula
        Assert.That(double.Parse(table.Rows[0][1].ToString()), Is.EqualTo(1.02));
        Assert.That(double.Parse(table.Rows[2][1].ToString()), Is.EqualTo(4.08));
        Assert.That(double.Parse(table.Rows[29][1].ToString()), Is.EqualTo(547608330.24));

        // Date + Formula
        Assert.That(DateTime.Parse(table.Rows[0][2].ToString()).ToShortDateString(), Is.EqualTo(new DateTime(2009, 5, 11).ToShortDateString()));
        Assert.That(DateTime.Parse(table.Rows[29][2].ToString()).ToShortDateString(), Is.EqualTo(new DateTime(2009, 11, 30).ToShortDateString()));
    }

    [Test]
    public void MultiSheetTest()
    {
        using IExcelDataReader excelReader = OpenReader("MultiSheet");
        DataSet result = excelReader.AsDataSet();

        Assert.That(result.Tables.Count, Is.EqualTo(3));

        Assert.That(result.Tables["Sheet1"].Columns.Count, Is.EqualTo(4));
        Assert.That(result.Tables["Sheet1"].Rows.Count, Is.EqualTo(12));
        Assert.That(result.Tables["Sheet2"].Columns.Count, Is.EqualTo(4));
        Assert.That(result.Tables["Sheet2"].Rows.Count, Is.EqualTo(12));
        Assert.That(result.Tables["Sheet3"].Columns.Count, Is.EqualTo(2));
        Assert.That(result.Tables["Sheet3"].Rows.Count, Is.EqualTo(5));

        Assert.That(result.Tables["Sheet2"].Rows[11][0].ToString(), Is.EqualTo("1"));
        Assert.That(result.Tables["Sheet1"].Rows[11][3].ToString(), Is.EqualTo("2"));
        Assert.That(result.Tables["Sheet3"].Rows[4][1].ToString(), Is.EqualTo("3"));
    }

    [Test]
    public void DataReaderNextResultTest()
    {
        using IExcelDataReader r = OpenReader("MultiSheet");
        Assert.That(r.ResultsCount, Is.EqualTo(3));

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

        Assert.That(table.Rows.Count, Is.EqualTo(12));
        Assert.That(r.RowCount, Is.EqualTo(12));
        Assert.That(fieldCount, Is.EqualTo(4));
        Assert.That(table.Rows[11][3], Is.EqualTo(1));

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

        Assert.That(table.Rows.Count, Is.EqualTo(12));
        Assert.That(r.RowCount, Is.EqualTo(12));
        Assert.That(fieldCount, Is.EqualTo(4));
        Assert.That(table.Rows[11][3], Is.EqualTo(2));

        r.NextResult();
        table.Rows.Clear();

        while (r.Read())
        {
            fieldCount = r.FieldCount;
            table.Rows.Add(
                Convert.ToInt32(r.GetValue(0)),
                Convert.ToInt32(r.GetValue(1)));
        }

        Assert.That(table.Rows.Count, Is.EqualTo(5));
        Assert.That(r.RowCount, Is.EqualTo(5));
        Assert.That(fieldCount, Is.EqualTo(2));
        Assert.That(table.Rows[4][1], Is.EqualTo(3));

        Assert.That(r.NextResult(), Is.EqualTo(false));
    }

    [Test]
    public void UnicodeCharsTest()
    {
        using IExcelDataReader excelReader = OpenReader("UnicodeChars");
        DataTable result = excelReader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(3));
        Assert.That(result.Columns.Count, Is.EqualTo(8));
        Assert.That(result.Rows[1][0].ToString(), Is.EqualTo("\u00e9\u0417"));
    }

    [Test]
    public void Issue29_ReadSheetStatesReadsCorrectly()
    {
        using IExcelDataReader excelReader = OpenReader("ExcelDataset");
        Assert.That(excelReader.VisibleState, Is.EqualTo("hidden"));

        excelReader.NextResult();
        Assert.That(excelReader.VisibleState, Is.EqualTo("visible"));

        excelReader.NextResult();
        Assert.That(excelReader.VisibleState, Is.EqualTo("veryhidden"));
    }

    [Test]
    public void Issue29_AsDataSetProvidesCorrectSheetState()
    {
        using IExcelDataReader reader = OpenReader("ExcelDataset");
        var dataSet = reader.AsDataSet();

        Assert.That(dataSet != null, Is.True);
        Assert.That(dataSet.Tables.Count, Is.EqualTo(3));
        Assert.That(dataSet.Tables[0].ExtendedProperties["visiblestate"], Is.EqualTo("hidden"));
        Assert.That(dataSet.Tables[1].ExtendedProperties["visiblestate"], Is.EqualTo("visible"));
        Assert.That(dataSet.Tables[2].ExtendedProperties["visiblestate"], Is.EqualTo("veryhidden"));
    }

    [Test]
    public void Issue389_FilterSheetByVisibility()
    {
        using IExcelDataReader excelReader = OpenReader("ExcelDataset");
        var result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
        {
            FilterSheet = (r, index) => r.VisibleState == "visible"
        });

        Assert.That(result.Tables.Count, Is.EqualTo(1));
    }

    [Test]
    public void TestNumDoubleDateBoolString()
    {
        using IExcelDataReader excelReader = OpenReader("NumDoubleDateBoolString");
        DataSet dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(30));
        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(6));

        Assert.That(int.Parse(dataSet.Tables[0].Rows[0][0].ToString()), Is.EqualTo(1));
        Assert.That(int.Parse(dataSet.Tables[0].Rows[29][0].ToString()), Is.EqualTo(1346269));

        // bool        
        Assert.That(dataSet.Tables[0].Rows[22][3].ToString(), Is.Not.Null);
        Assert.That(dataSet.Tables[0].Rows[22][3], Is.EqualTo(true));

        // double + Formula
        Assert.That(double.Parse(dataSet.Tables[0].Rows[0][1].ToString()), Is.EqualTo(1.02));
        Assert.That(double.Parse(dataSet.Tables[0].Rows[2][1].ToString()), Is.EqualTo(4.08));
        Assert.That(double.Parse(dataSet.Tables[0].Rows[29][1].ToString()), Is.EqualTo(547608330.24));

        // Date + Formula
        string s = dataSet.Tables[0].Rows[0][2].ToString();
        Assert.That(DateTime.Parse(s), Is.EqualTo(new DateTime(2009, 5, 11)));
        s = dataSet.Tables[0].Rows[29][2].ToString();
        Assert.That(DateTime.Parse(s), Is.EqualTo(new DateTime(2009, 11, 30)));

        // Custom Date Time + Formula
        s = dataSet.Tables[0].Rows[0][5].ToString();
        Assert.That(DateTime.Parse(s), Is.EqualTo(new DateTime(2009, 5, 7, 11, 1, 2)));
        s = dataSet.Tables[0].Rows[1][5].ToString();
        Assert.That(DateTime.Parse(s), Is.EqualTo(new DateTime(2009, 5, 8, 11, 1, 2)));

        // DBNull value
        Assert.That(dataSet.Tables[0].Rows[1][4], Is.EqualTo(DBNull.Value));
    }

    [Test]
    public void Issue160_FilterRow()
    {
        // Check there are four rows with data, including empty and hidden rows
        using (var reader = OpenReader("CollapsedHide"))
        {
            var dataSet = reader.AsDataSet();

            Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(4));
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

            Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(2));
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

            Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(1));
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
    public void Issue300_FilterColumn()
    {
        // Check there are two columns with data
        using (var reader = OpenReader("CollapsedHide"))
        {
            var dataSet = reader.AsDataSet();

            Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(2));
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

            Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(1));
        }
    }

    [Test]
    public void Issue483_CellErrorEmptyRow()
    {
        // Check there are four rows with no errors and no NREs
        using var reader = OpenReader("CollapsedHide");
        reader.Read();
        Assert.That(reader.GetCellError(0), Is.EqualTo(null));
        Assert.That(reader.GetCellError(1), Is.EqualTo(null));

        reader.Read();
        Assert.That(reader.GetCellError(0), Is.EqualTo(null));
        Assert.That(reader.GetCellError(1), Is.EqualTo(null));

        reader.Read();
        Assert.That(reader.GetCellError(0), Is.EqualTo(null));
        Assert.That(reader.GetCellError(1), Is.EqualTo(null));

        reader.Read();
        Assert.That(reader.GetCellError(0), Is.EqualTo(null));
        Assert.That(reader.GetCellError(1), Is.EqualTo(null));
    }

    [Test]
    public void Issue532_TrimEmptyColumns()
    {
        using var reader = OpenReader("Issue532");
        while (reader.Read())
        {
            Assert.That(reader.FieldCount, Is.EqualTo(3));
        }
    }

    [Test]
    public void Issue694_ExcelTimeFormatTimeSpan()
    {
        using var reader = OpenReader("Issue694_TimeSpan");
        reader.Read();
        reader.Read();
        Assert.That(TimeSpan.Parse(reader[1].ToString()), Is.EqualTo(TimeSpan.Parse("-13:57")));
        reader.Read();
        Assert.That(TimeSpan.Parse(reader[1].ToString()), Is.EqualTo(TimeSpan.Parse("-00:11")));
        reader.Read();
        Assert.That(TimeSpan.Parse(reader[1].ToString()), Is.EqualTo(TimeSpan.Parse("00:11")));
    }

    [Test]
    public void Issue694_ExcelTimeFormatTimeSpanFormulaInvalidResult()
    {
        using var reader = OpenReader("Issue694_TimeSpanFormula");
        reader.Read();
        Assert.Multiple(() =>
        {
            Assert.That(reader[0].ToString(), Is.EqualTo("A"));
            Assert.That(reader[1].ToString(), Is.EqualTo("P"));
        });
    }

    [Test]
    public void Issue574_VerticalAlignment()
    {
        using var reader = OpenReader("Issue574");
        reader.Read();

        Assert.That(reader.GetCellStyle(0).VerticalAlignment, Is.EqualTo(VerticalAlignment.Top));
        Assert.That(reader.GetCellStyle(1).VerticalAlignment, Is.EqualTo(VerticalAlignment.Center));
        Assert.That(reader.GetCellStyle(2).VerticalAlignment, Is.EqualTo(VerticalAlignment.Bottom));
    }

    [Test]
    public void Issue618_SinglePassMode_RowCountThrows()
    {
        using var reader = OpenReader(OpenStream("10x10"), new ExcelReaderConfiguration { SinglePassMode = true });
        Assert.Throws<InvalidOperationException>(() => _ = reader.RowCount);
        reader.Read();
        Assert.Throws<InvalidOperationException>(() => _ = reader.RowCount);
    }

    [Test]
    public void Issue618_SinglePassMode_AsDataSet()
    {
        using var reader = OpenReader(OpenStream("10x10"), new ExcelReaderConfiguration { SinglePassMode = true });
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(10));
        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(10));
        Assert.That(dataSet.Tables[0].Rows[1][0], Is.EqualTo("10x10"));
        Assert.That(dataSet.Tables[0].Rows[9][9], Is.EqualTo("10x27"));
    }

    [Test]
    public void Issue618_SinglePassMode_FieldCountGrows()
    {
        using var reader = OpenReader(OpenStream("10x10"), new ExcelReaderConfiguration { SinglePassMode = true });
        Assert.That(reader.FieldCount, Is.Zero);
        reader.Read();
        Assert.That(reader.FieldCount, Is.EqualTo(9));
        while (reader.Read())
        {            
        }
        
        Assert.That(reader.FieldCount, Is.GreaterThanOrEqualTo(10));
    }

    public void Issue541_BuiltinFormat55IsDate()
    {
        using var reader = OpenReader("Issue541");
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetValue(0), Is.EqualTo(new DateTime(2021, 1, 15)));
    }

    public void AsDataSetTestFillEmptyCellsInMergedRangeNotUseHeaderRow()
    {
        using IExcelDataReader excelReader = OpenReader("MergedCell");
        DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration
            {
                FillMergedCellsValue = true
            }
        });

        Assert.That(result != null, Is.True);
        Assert.That(result.Tables.Count, Is.EqualTo(1));
        Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(7));
        Assert.That(result.Tables[0].Columns.Count, Is.EqualTo(3));

        Assert.That(result.Tables[0].Rows[1][0], Is.EqualTo("Merge Cell 1"));
        Assert.That(result.Tables[0].Rows[2][0], Is.EqualTo("Merge Cell 1"));
        Assert.That(result.Tables[0].Rows[1][1], Is.EqualTo("Merge Cell 1"));
        Assert.That(result.Tables[0].Rows[2][1], Is.EqualTo("Merge Cell 1"));
        Assert.That(result.Tables[0].Rows[5][2], Is.EqualTo("Merge Cell 2"));
        Assert.That(result.Tables[0].Rows[4][0], Is.EqualTo("Merge Cell 3"));
        Assert.That(result.Tables[0].Rows[6][1], Is.EqualTo("Merge Cell 4"));
        Assert.That(result.Tables[0].Rows[6][2], Is.EqualTo("Merge Cell 4"));
    }

    [Test]
    public void AsDataSetTestFillEmptyCellsInMergedRangeUseHeaderRow()
    {
        using IExcelDataReader excelReader = OpenReader("MergedCell");
        DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration
            {
                UseHeaderRow = true,
                FillMergedCellsValue = true
            }
        });

        Assert.That(result != null, Is.True);
        Assert.That(result.Tables.Count, Is.EqualTo(1));
        Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(6));
        Assert.That(result.Tables[0].Columns.Count, Is.EqualTo(3));

        Assert.That(result.Tables[0].Rows[0][0], Is.EqualTo("Merge Cell 1"));
        Assert.That(result.Tables[0].Rows[1][0], Is.EqualTo("Merge Cell 1"));
        Assert.That(result.Tables[0].Rows[0][1], Is.EqualTo("Merge Cell 1"));
        Assert.That(result.Tables[0].Rows[1][1], Is.EqualTo("Merge Cell 1"));
        Assert.That(result.Tables[0].Rows[4][2], Is.EqualTo("Merge Cell 2"));
        Assert.That(result.Tables[0].Rows[3][0], Is.EqualTo("Merge Cell 3"));
        Assert.That(result.Tables[0].Rows[5][1], Is.EqualTo("Merge Cell 4"));
        Assert.That(result.Tables[0].Rows[5][2], Is.EqualTo("Merge Cell 4"));
    }

    [Test]
    public void AsDataSetFillMergedCellsValueWithNoMergeCells()
    {
        using IExcelDataReader excelReader = OpenReader("10x10");
        DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration
            {
                UseHeaderRow = true,
                FillMergedCellsValue = true
            }
        });

        Assert.That(result, Is.Not.Null);
        Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(9));
        Assert.That(result.Tables[0].Columns.Count, Is.EqualTo(10));
    }
    
    protected IExcelDataReader OpenReader(string name)
    {
        return OpenReader(OpenStream(name));
    }

    protected abstract Stream OpenStream(string name);

    protected abstract IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null);
}
