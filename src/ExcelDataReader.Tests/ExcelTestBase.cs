using System.Data;
using System.Globalization;

namespace ExcelDataReader.Tests;

[Parallelizable]
public abstract class ExcelTestBase
{
    protected abstract DateTime Issue82_TodayDate { get; }

    protected virtual bool SupportsCodeName => true;

    [TestCase(false)]
    [TestCase(true)]
    public void IssueFileLock5161_NextResultCanAbandonPartiallyReadSheets(bool singlePass)
    {
        using var reader = OpenReader(OpenFixtureStream("MultiSheet"), new ExcelReaderConfiguration { SinglePassMode = singlePass });
        int sheets = 0;
        do
        {
            Assert.That(reader.Read(), Is.True);
            sheets++;
        }
        while (reader.NextResult());
        Assert.That(sheets, Is.EqualTo(reader.ResultsCount));
    }

    [Test]
    public void Issue486_TransformValue()
    {
        using var reader = OpenReader("Issue486");
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            // DataSet converts boxed enum values to integers when column types are inferred.
            UseColumnDataType = false,
            ConfigureDataTable = _ => new ExcelDataTableConfiguration
            {
                UseHeaderRow = true,
                TransformValue = (transformReader, n, value) =>
                {
                    var error = transformReader.GetCellError(n);
                    if (error != null)
                        return error;
                    return value;
                }
            }
        });
        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo(CellError.REF));
        Assert.That(dataSet.Tables[0].Rows[0][1], Is.EqualTo(CellError.REF));
        Assert.That(dataSet.Tables[0].Rows[1][0], Is.EqualTo(CellError.NAME));
        Assert.That(dataSet.Tables[0].Rows[1][1], Is.EqualTo(CellError.NAME));
    }

    [Test]
    public void Issue518_MultipleHeaderRows()
    {
        using var reader = OpenReader("Issue518");
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration
            {
                UseHeaderRow = true,
                ReadHeader = self =>
                {
                    List<string> headerNames = [];
                    for (int i = 0; i < self.FieldCount; i++)
                        headerNames.Add(Convert.ToString(self.GetValue(i)));
                    Assert.That(self.Read(), Is.True);
                    var result = new Dictionary<int, string>(self.FieldCount);
                    for (int i = 0; i < self.FieldCount; i++)
                    {
                        string first = headerNames[i];
                        string second = Convert.ToString(self.GetValue(i));
                        string name = first.Length == 0 ? second : second.Length == 0 ? first : first + " " + second;
                        result.Add(i, string.IsNullOrEmpty(name) ? "Column" + i : name);
                    }

                    return result;
                }
            }
        });
        var columns = dataSet.Tables[0].Columns;
        Assert.That(columns[0].ColumnName, Is.EqualTo("ColName1 A"));
        Assert.That(columns[1].ColumnName, Is.EqualTo("ColName1 B"));
        Assert.That(columns[2].ColumnName, Is.EqualTo("ColName2 B"));
        Assert.That(columns[3].ColumnName, Is.EqualTo("FirstOnly"));
        Assert.That(columns[4].ColumnName, Is.EqualTo("SecondOnly"));
        Assert.That(columns[5].ColumnName, Is.EqualTo("Another One"));
    }

    [Test]
    public void Issue240_ExceptionBeforeRead()
    {
        using var reader = OpenReader("10x10");
        var exception = Assert.Throws<InvalidOperationException>(() => reader.GetString(0));
        Assert.That(exception.Message, Is.EqualTo("No data exists for the row/column."));
    }

    [Test]
    public void Issue525_StringsAndMissingCellsReturnCorrectValues()
    {
        using var reader = OpenReader("10x10");
        Assert.That(reader.Read(), Is.True);
        string first = reader.GetString(0);
        Assert.That(first, Is.EqualTo("col1"));
        Assert.That(reader.GetString(4), Is.EqualTo("col5"));
        Assert.That(reader.IsDBNull(9), Is.True);
        Assert.Throws<InvalidCastException>(() => reader.GetString(9));
        Assert.That(reader.GetString(0), Is.SameAs(first));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("10x10"));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Issue734_DateTimeCellsReturnDateTime(bool singlePass)
    {
        using var reader = OpenReader(OpenFixtureStream("NumDoubleDateBoolString"), new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Read(), Is.True);
        var value = reader.GetValue(5);
        Assert.That(value, Is.TypeOf<DateTime>());
        Assert.That(reader.GetFieldType(5), Is.EqualTo(typeof(DateTime)));
        Assert.That(reader.GetDateTime(5), Is.EqualTo((DateTime)value));
    }

    [Test]
    public void Issue578()
    {
        using var reader = OpenReader("Issue578");
        Assert.That(reader.Read(), Is.True);
        var values = new object[reader.FieldCount];
        Assert.That(reader.GetValues(values), Is.EqualTo(reader.FieldCount));
        var values2 = new object[reader.FieldCount + 1];
        Assert.That(reader.GetValues(values2), Is.EqualTo(reader.FieldCount));
        var values3 = new object[reader.FieldCount - 1];
        Assert.That(reader.GetValues(values3), Is.EqualTo(values3.Length));
        Assert.That(values, Is.EqualTo(new object[] { 1, 2, 3, 4, 5 }));
        Assert.That(values2, Is.EqualTo(new object[] { 1, 2, 3, 4, 5, null }));
        Assert.That(values3, Is.EqualTo(new object[] { 1, 2, 3, 4 }));
    }

    [Test]
    public void Issue642_ActiveSheet()
    {
        using var reader = OpenReader("Issue642");
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            FilterSheet = (tableReader, sheetIndex) => tableReader.IsActiveSheet
        });
        Assert.That(reader.ActiveSheet, Is.EqualTo(5));
        Assert.That(dataSet.Tables.Count, Is.EqualTo(1));
        Assert.That(dataSet.Tables[0].TableName, Is.EqualTo("List6"));
    }

    [Test]
    public void Issue642_ActiveSheet_SingleWorksheet()
    {
        using var reader = OpenReader("Issue642_Onesheet");
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            FilterSheet = (tableReader, sheetIndex) => tableReader.IsActiveSheet
        });
        Assert.That(reader.ActiveSheet, Is.EqualTo(0));
        Assert.That(dataSet.Tables.Count, Is.EqualTo(1));
        Assert.That(dataSet.Tables[0].TableName, Is.EqualTo("List1"));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void LazyCommentsPreserveBufferedRowsAndReaderLifecycle(bool singlePass)
    {
        const string text = "first & <second>\n\u03bb \ud83d\ude00";
        using var source = OpenFixtureStream("CellComments");
        using var reader = OpenReader(source, new ExcelReaderConfiguration { SinglePassMode = singlePass, LeaveOpen = true });
        Assert.Throws<InvalidOperationException>(() => reader.GetComment(0));
        CellCommentThread retained = null;
        for (int row = 0; row < 700; row++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(0), Is.EqualTo(row));
            if (row == 13)
            {
                retained = reader.GetComment(0);
                Assert.That(retained, Is.Not.Null);
                Assert.That(retained.Messages.Count, Is.EqualTo(1));
                Assert.That(retained.Messages[0].Text, Is.EqualTo(text));
                Assert.That(reader.GetComment(0), Is.SameAs(retained));
                Assert.That(reader.GetDouble(0), Is.EqualTo(row));
                Assert.Throws<IndexOutOfRangeException>(() => reader.GetComment(-1));
                Assert.Throws<IndexOutOfRangeException>(() => reader.GetComment(reader.FieldCount));
            }
            else
            {
                Assert.That(reader.GetComment(0), Is.Null);
            }
        }

        Assert.That(reader.Read(), Is.False);
        Assert.Throws<InvalidOperationException>(() => reader.GetComment(0));
        reader.Reset();
        for (int attempt = 0; attempt < 2; attempt++)
        {
            for (int row = 0; row <= 13; row++)
                Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetComment(0), Is.SameAs(retained));
            reader.ResetWorksheet();
        }

        reader.Close();
        Assert.That(source.CanRead, Is.True);
        Assert.Throws<InvalidOperationException>(() => reader.GetComment(0));
        Assert.That(retained.Messages[0].Text, Is.EqualTo(text));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void ReturnedStylesRemainStableAcrossReadingResetAndDisposal(bool singlePass)
    {
        using var reader = OpenReader(OpenFixtureStream("NumDoubleDateBoolString"), new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Read(), Is.True);
        var style = reader.GetCellStyle(0);
        var expected = (style.FontIndex, style.NumberFormatIndex, style.IndentLevel,
            style.HorizontalAlignment, style.VerticalAlignment, style.Hidden, style.Locked);
        Assert.That(reader.GetCellStyle(0), Is.SameAs(style));
        while (reader.Read())
        {
            for (int column = 0; column < reader.FieldCount; column++)
                _ = reader.GetCellStyle(column);
        }

        reader.Reset();
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetCellStyle(0), Is.SameAs(style));
        reader.Dispose();
        var actual = (style.FontIndex, style.NumberFormatIndex, style.IndentLevel,
            style.HorizontalAlignment, style.VerticalAlignment, style.Hidden, style.Locked);
        Assert.That(actual, Is.EqualTo(expected));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void CachedFormulaErrorsRemainSeparateFromNullValues(bool singlePass)
    {
        using var reader = OpenReader(OpenFixtureStream("Issue329_Error"), new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Read(), Is.True);
        reader.Reset();
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Read(), Is.True);
            for (int column = 0; column < 3; column++)
                Assert.That(reader.IsDBNull(column), Is.True);
            Assert.That(reader.GetCellError(0), Is.EqualTo(CellError.DIV0));
            Assert.That(reader.GetCellError(1), Is.EqualTo(CellError.NA));
            Assert.That(reader.GetCellError(2), Is.EqualTo(CellError.VALUE));
            while (reader.Read())
            {
            }

            reader.Reset();
        }
    }

    [TestCase(false)]
    [TestCase(true)]
    public void CachedFormulaStringsAndDatesSurviveReset(bool singlePass)
    {
        using var reader = OpenReader(OpenFixtureStream("EncodingFormulaDate1520"), new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Read(), Is.True);
        reader.Reset();
        for (int pass = 0; pass < 2; pass++)
        {
            for (int row = 0; row <= 8; row++)
            {
                Assert.That(reader.Read(), Is.True);
                if (row == 1)
                {
                    Assert.That(reader.GetString(0), Is.EqualTo("John test"));
                    Assert.That(reader.GetDateTime(1).Date, Is.EqualTo(new DateTime(2009, 5, 1)));
                }
                else if (row == 2)
                {
                    Assert.That(reader.GetString(0), Is.EqualTo("Simon Hodgetts"));
                    Assert.That(reader.GetDateTime(4).TimeOfDay, Is.EqualTo(TimeSpan.FromHours(11)));
                }
                else if (row is 7 or 8)
                {
                    Assert.That(reader.GetString(0), Is.EqualTo("librement réutilisable"));
                }
            }

            while (reader.Read())
            {
            }

            reader.Reset();
        }
    }

    [TestCase(false)]
    [TestCase(true)]
    public void PartialReadResetAndNextResultPreserveValuesAndLeaveOpen(bool singlePass)
    {
        using var stream = OpenFixtureStream("10x10000");
        using (var reader = OpenReader(stream, new ExcelReaderConfiguration { SinglePassMode = singlePass, LeaveOpen = true }))
        {
            Assert.That(reader.Read(), Is.True);
            object first = reader.GetValue(0);
            reader.Reset();
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetValue(0), Is.EqualTo(first));
            _ = reader.NextResult();
            reader.Reset();
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetValue(0), Is.EqualTo(first));
        }

        Assert.That(stream.CanRead, Is.True);
    }

    [Test]
    public void ResetWorksheetReplaysCurrentResultWithoutChangingNavigation(
        [Values("10x10", "Open")] string fixture, [Values(false, true)] bool singlePass)
    {
        using var source = OpenFixtureStream(fixture);
        var configuration = new ExcelReaderConfiguration { SinglePassMode = singlePass, LeaveOpen = true };
        using var reader = OpenReader(source, configuration);
        using var baseline = OpenReader(OpenFixtureStream(fixture), configuration);
        AssertResetWorksheetReplaysCurrentResult(reader, baseline, source);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void DocumentPropertiesAreAvailableDuringOpenAndSurviveReading(bool singlePass)
    {
        using var source = OpenStream("DocumentProperties");
        using var reader = OpenReader(source, new ExcelReaderConfiguration { SinglePassMode = singlePass, LeaveOpen = true });
        AssertDocumentPropertiesSurviveReading(reader, source);
    }

    [Test]
    public void CsvSeparatorIsNullForSpreadsheets()
    {
        using var reader = OpenReader("OldIssue11573_BlankValues");
        Assert.That(reader.CsvSeparator, Is.Null);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.CsvSeparator, Is.Null);
    }

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
        Assert.That(reader.GetColumnWidth(2), Is.EqualTo(15.140625).Within(0.01));
        Assert.That(reader.GetColumnWidth(3), Is.EqualTo(28.7109375).Within(0.01));

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
            var stream = OpenFixtureStream("10x10");
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
        if (!SupportsCodeName)
            Assert.Ignore($"{GetType().Name} does not support sheet code names.");

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
        Assert.That(reader.GetValue(0), Is.EqualTo(DBNull.Value));
        Assert.That(reader.GetCellError(0), Is.EqualTo(CellError.DIV0));

        Assert.That(reader.GetValue(1), Is.EqualTo(DBNull.Value));
        Assert.That(reader.GetCellError(1), Is.EqualTo(CellError.NA));

        Assert.That(reader.GetValue(2), Is.EqualTo(DBNull.Value));
        Assert.That(reader.GetCellError(2), Is.EqualTo(CellError.VALUE));

        Assert.That(reader.RowCount, Is.EqualTo(1));
    }

    [Test]
    public void Issue4031_NullColumn()
    {
        using IExcelDataReader excelReader = OpenReader("OldIssue4031_NullColumn");

        // DataSet dataSet = excelReader.AsDataSet(true);
        excelReader.Read();
        Assert.That(excelReader.GetValue(0), Is.EqualTo(DBNull.Value));
        Assert.That(excelReader.GetString(1), Is.EqualTo("a"));
        Assert.That(excelReader.GetString(2), Is.EqualTo("b"));
        Assert.That(excelReader.GetValue(3), Is.EqualTo(DBNull.Value));
        Assert.That(excelReader.GetString(4), Is.EqualTo("d"));

        excelReader.Read();
        Assert.That(excelReader.GetValue(0), Is.EqualTo(DBNull.Value));
        Assert.That(excelReader.GetValue(1), Is.EqualTo(DBNull.Value));
        Assert.That(excelReader.GetString(2), Is.EqualTo("Test"));
        Assert.That(excelReader.GetValue(3), Is.EqualTo(DBNull.Value));
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
    public void Issue443_DepthAlwaysZero()
    {
        using IExcelDataReader reader = OpenReader("MultiSheet");
        Assert.That(reader.Depth, Is.Zero);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.Depth, Is.Zero);

        while (reader.Read())
        {
            Assert.That(reader.Depth, Is.Zero);
        }

        Assert.That(reader.NextResult(), Is.True);
        Assert.That(reader.Depth, Is.Zero);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.Depth, Is.Zero);
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
                if (reader.GetValue(i) is not DBNull)
                    return false;
            }

            return true;
        }

        static bool IsEmptyRow(IExcelDataReader reader)
        {
            for (var i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetValue(i) is not DBNull)
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
        using var reader = OpenReader(OpenFixtureStream("10x10"), new ExcelReaderConfiguration { SinglePassMode = true });
        Assert.Throws<InvalidOperationException>(() => _ = reader.RowCount);
        reader.Read();
        Assert.Throws<InvalidOperationException>(() => _ = reader.RowCount);
    }

    [Test]
    public void Issue618_SinglePassMode_AsDataSet()
    {
        using var reader = OpenReader(OpenFixtureStream("10x10"), new ExcelReaderConfiguration { SinglePassMode = true });
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(10));
        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(10));
        Assert.That(dataSet.Tables[0].Rows[1][0], Is.EqualTo("10x10"));
        Assert.That(dataSet.Tables[0].Rows[9][9], Is.EqualTo("10x27"));
    }

    [Test]
    public void Issue618_SinglePassMode_FieldCountGrows()
    {
        using var reader = OpenReader(OpenFixtureStream("10x10"), new ExcelReaderConfiguration { SinglePassMode = true });
        Assert.That(reader.FieldCount, Is.Zero);
        reader.Read();
        Assert.That(reader.FieldCount, Is.EqualTo(9));
        while (reader.Read())
        {
        }

        Assert.That(reader.FieldCount, Is.GreaterThanOrEqualTo(10));
    }

    [Test]
    public void Issue541_BuiltinFormat55IsDate()
    {
        using var reader = OpenReader("Issue541");
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetValue(0), Is.EqualTo(new DateTime(2021, 1, 15)));
    }

    [Test]
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

    internal static void AssertResetWorksheetReplaysCurrentResult(IExcelDataReader reader, IExcelDataReader baseline, Stream source)
    {
        do
        {
            string name = baseline.Name;
            int activeSheet = baseline.ActiveSheet;
            bool isActiveSheet = baseline.IsActiveSheet;
            var rows = new List<(object[] Values, double Height)>();
            while (baseline.Read())
            {
                var values = new object[baseline.FieldCount];
                baseline.GetValues(values);
                rows.Add((values, baseline.RowHeight));
            }

            reader.ResetWorksheet();
            _ = reader.Read();
            for (int pass = 0; pass < 2; pass++)
            {
                reader.ResetWorksheet();
                Assert.That(reader.Name, Is.EqualTo(name));
                Assert.That(reader.ActiveSheet, Is.EqualTo(activeSheet));
                Assert.That(reader.IsActiveSheet, Is.EqualTo(isActiveSheet));
                Assert.That(reader.RowHeight, Is.Zero);
                Assert.Throws<InvalidOperationException>(() => reader.GetValue(0));
                foreach (var row in rows)
                {
                    Assert.That(reader.Read(), Is.True);
                    var values = new object[reader.FieldCount];
                    reader.GetValues(values);
                    Assert.That(values, Is.EqualTo(row.Values));
                    Assert.That(reader.RowHeight, Is.EqualTo(row.Height));
                }

                Assert.That(reader.Read(), Is.False);
            }

            bool next = baseline.NextResult();
            Assert.That(reader.NextResult(), Is.EqualTo(next));
            if (!next)
                break;
        }
        while (true);

        Assert.Throws<InvalidOperationException>(reader.ResetWorksheet);
        reader.Close();
        Assert.Throws<InvalidOperationException>(reader.ResetWorksheet);
        Assert.That(source.CanRead, Is.True);
    }

    internal static void AssertDocumentPropertiesSurviveReading(IExcelDataReader reader, Stream source)
    {
        var properties = reader.DocumentProperties;
        Assert.That(properties.Title, Is.EqualTo("Workbook title"));
        Assert.That(properties.Subject, Is.EqualTo("Workbook subject"));
        Assert.That(properties.Keywords, Is.EqualTo("excel; test; keywords"));
        Assert.That(properties.Description, Is.EqualTo("Description & <text>\nSecond line"));

        // Excel's RemovePersonalInformation option removes both author fields on save.
        Assert.That(properties.Author, Is.Null.Or.Empty);
        Assert.That(properties.LastModifiedBy, Is.Null.Or.Empty);
        Assert.That(properties.Created, Is.EqualTo(new DateTime(2020, 1, 2, 2, 4, 5, DateTimeKind.Utc)));
        Assert.That(properties.LastPrinted, Is.EqualTo(new DateTime(2020, 5, 6, 5, 8, 8, DateTimeKind.Utc)));
        Assert.That(properties.Modified, Is.GreaterThan(properties.Created));
        Assert.That(properties.Created!.Value.Kind, Is.EqualTo(DateTimeKind.Utc));
        Assert.That(properties.Modified!.Value.Kind, Is.EqualTo(DateTimeKind.Utc));
        Assert.That(properties.LastPrinted!.Value.Kind, Is.EqualTo(DateTimeKind.Utc));
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Name, Is.EqualTo("Data"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo("Document properties fixture"));
            Assert.That(reader.GetDouble(1), Is.EqualTo(42));
            reader.ResetWorksheet();
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(1), Is.EqualTo(42));
            Assert.That(reader.Read(), Is.False);
            Assert.That(reader.NextResult(), Is.True);
            Assert.That(reader.Name, Is.EqualTo("Other"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo("Second sheet"));
            Assert.That(reader.Read(), Is.False);
            Assert.That(reader.NextResult(), Is.False);
            Assert.That(reader.DocumentProperties, Is.SameAs(properties));
            reader.Reset();
        }

        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables.Count, Is.EqualTo(2));
        Assert.That(dataSet.Tables[0].Rows[0][1], Is.EqualTo(42));
        reader.Close();
        Assert.That(source.CanRead, Is.True);
        Assert.That(reader.DocumentProperties, Is.SameAs(properties));
        Assert.That(properties.Title, Is.EqualTo("Workbook title"));
    }

    protected IExcelDataReader OpenReader(string name)
    {
        return OpenReader(OpenFixtureStream(name));
    }

    protected Stream OpenFixtureStream(string name)
    {
        if (!IsFixtureSupported(name))
            Assert.Ignore($"Fixture '{name}' is not supported by {GetType().Name}.");

        return OpenStream(name);
    }

    protected virtual bool IsFixtureSupported(string name) => true;

    protected abstract Stream OpenStream(string name);

    protected abstract IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null);
}