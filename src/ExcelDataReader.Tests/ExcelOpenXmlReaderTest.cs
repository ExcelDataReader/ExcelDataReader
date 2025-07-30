using System.Data;

namespace ExcelDataReader.Tests;

public class ExcelOpenXmlReaderTest : ExcelOpenXmlReaderBase
{
    /// <inheritdoc />
    protected override DateTime GitIssue82TodayDate => new(2013, 4, 19);

    [Test]
    public void FailTest()
    {
        var expectedException = typeof(Exceptions.HeaderException);

        var exception = Assert.Throws(expectedException, () =>
            {
                using (ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestFail_Binary.xls")))
                {
                }
            });

        Assert.That(exception.Message, Is.EqualTo("Invalid file signature."));
    }

    [Test]
    public void Issue4145()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_4145.xlsx"));
        Assert.DoesNotThrow(() => excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration));

        while (excelReader.Read())
        {
        }
    }

    [Test]
    public void IssueFileLock5161()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestMultiSheet.xlsx"));
        
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

    [Test]
    public void Issue11522OpenXml()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11522_OpenXml.xlsx"));
        DataSet result = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

        Assert.That(result.Tables[0].Columns.Count, Is.EqualTo(11));
        Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(1));
        Assert.That(result.Tables[0].Rows[0][0], Is.EqualTo("TestNewButton"));
        Assert.That(result.Tables[0].Rows[0][1], Is.EqualTo("677"));
        Assert.That(result.Tables[0].Rows[0][2], Is.EqualTo("u77"));
        Assert.That(result.Tables[0].Rows[0][3], Is.EqualTo("u766"));
        Assert.That(result.Tables[0].Rows[0][4], Is.EqualTo("y66"));
        Assert.That(result.Tables[0].Rows[0][5], Is.EqualTo("F"));
        Assert.That(result.Tables[0].Rows[0][6], Is.EqualTo(DBNull.Value));
        Assert.That(result.Tables[0].Rows[0][7], Is.EqualTo(DBNull.Value));
        Assert.That(result.Tables[0].Rows[0][8], Is.EqualTo(DBNull.Value));
        Assert.That(result.Tables[0].Rows[0][9], Is.EqualTo(DBNull.Value));
        Assert.That(result.Tables[0].Rows[0][10], Is.EqualTo(DBNull.Value));
    }

    /*
#if !LEGACY
            [Test]
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

    [Test]
    public void TestGoogleSourced()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_googlesourced.xlsx"));
        var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

        Assert.That(dataSet.Tables[0].Rows[0][0].ToString(), Is.EqualTo("9583638582"));
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(4));
        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(6));
    }

    [Test]
    public void TestIssue12667GoogleExportMissingColumns()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_12667_GoogleExport_MissingColumns.xlsx"));
        var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(6)); // 6 with data + 1 that is present but no data in it
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(0));
    }

    [Test]
    public void IssueGit142()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_Issue_142.xlsx"));
        var dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(4));
    }

    /// <summary>
    /// Sheet has no [dimension] and/or no [cols].
    /// Sheet has no [styles].
    /// Each row [row] has no "r" attribute.
    /// Each cell [c] has no "r" attribute.
    /// </summary>
    [Test]
    public void IssueNoStylesNoRAttribute()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_NoStyles_NoRAttribute.xlsx"));
        DataSet result = excelReader.AsDataSet();

        Assert.That(result.Tables.Count, Is.GreaterThan(0));
        Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(39));
        Assert.That(result.Tables[0].Columns.Count, Is.EqualTo(18));
        Assert.That(result.Tables[0].Rows[4][4].ToString(), Is.EqualTo("ROW NUMBER 5"));

        excelReader.Close();
    }

    [Test]
    public void NoDimensionOrCellReferenceAttribute()
    {
        // 20170306_Daily Package GPR 250 Index EUR Overview.xlsx
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("NoDimensionOrCellReferenceAttribute.xlsx"));
        DataSet result = excelReader.AsDataSet();
        Assert.That(result.Tables.Count, Is.EqualTo(2));
        Assert.That(result.Tables[0].Columns.Count, Is.EqualTo(8), "Sheet0 Columns");
        Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(7), "Sheet0 Rows");

        Assert.That(result.Tables[1].Columns.Count, Is.EqualTo(8), "Sheet1 Columns");
        Assert.That(result.Tables[1].Rows.Count, Is.EqualTo(20), "Sheet1 Rows");
    }

    [Test]
    public void LowerCaseReferenceAttribute()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("LowerCaseReferenceAttribute.xlsx"));
        DataSet result = excelReader.AsDataSet();
        Assert.That(result.Tables[0].Rows[2][4], Is.EqualTo("E3"), "Sheet1 Cell E3");
    }

    [Test]
    public void CellValueIso8601Date()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_221.xlsx"));
        DataSet result = excelReader.AsDataSet();
        Assert.That(result.Tables[0].Rows[0][0], Is.EqualTo(new DateTime(2017, 3, 16)));
    }

    [Test]
    public void CellFormat49()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Format49_@.xlsx"));
        DataSet result = excelReader.AsDataSet();

        // ExcelDataReader used to convert numbers formatted with NumFmtId=49/@ to culture-specific strings.
        // This behaviour changed in v3 to return the original value:
        // Assert.That(result.Tables[0].Rows[0].ItemArray, Is.EqualTo(new[] { "2010-05-05", "1.1", "2,2", "123", "2,2" }));
        Assert.That(result.Tables[0].Rows[0].ItemArray, Is.EqualTo(new object[] { "2010-05-05", "1.1", 2.2000000000000002D, 123.0D, "2,2" }));
    }

    [Test]
    public void FullWidthSpace()
    {
        using var reader = OpenReader("Test_FullWidthSpace");
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].Rows[0].ItemArray[0], Is.EqualTo("\u3000\u3000text\u3000\u3000"));
        Assert.That(dataSet.Tables[0].Rows[1].ItemArray[0], Is.EqualTo("text\u3000"));
        Assert.That(dataSet.Tables[0].Rows[2].ItemArray[0], Is.EqualTo("\u3000text"));
    }
 
    [Test]
    public void GitIssue97()
    {
        // fillreport.xlsx was generated by a third party and uses badly formatted cell references with only numerals.
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("fillreport.xlsx"));
        DataSet result = excelReader.AsDataSet();
        Assert.That(result.Tables.Count, Is.EqualTo(1));
        Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(20));
        Assert.That(result.Tables[0].Columns.Count, Is.EqualTo(10));
        Assert.That(result.Tables[0].Rows[1][0], Is.EqualTo("Account Number"));
        Assert.That(result.Tables[0].Rows[1][1], Is.EqualTo("Trader"));
    }

    [Test]
    public void GitIssue68NullSheetPath()
    {
        using var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_68_NullSheetPath.xlsm"));
        DataSet result = excelReader.AsDataSet();
        Assert.That(result.Tables[0].Columns.Count, Is.EqualTo(2));
        Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(1));
    }

    [Test]
    public void GitIssue53CachedFormulaStringType()
    {
        using var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_53_Cached_Formula_String_Type.xlsx"));
        var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

        // Ensure that parseable, numeric cached formula values are read as a double
        Assert.That(dataSet.Tables[0].Rows[0][2], Is.InstanceOf<double>());
        Assert.That(dataSet.Tables[0].Rows[0][2], Is.EqualTo(3D));

        // Ensure that non-parseable, non-numeric cached formula values are read as a string
        Assert.That(dataSet.Tables[0].Rows[1][2], Is.InstanceOf<string>());
        Assert.That(dataSet.Tables[0].Rows[1][2], Is.EqualTo("AB"));

        // Ensure that parseable, non-numeric cached formula values are read as a string
        Assert.That(dataSet.Tables[0].Rows[2][2], Is.InstanceOf<string>());
        Assert.That(dataSet.Tables[0].Rows[2][2], Is.EqualTo("1,"));
    }

    [Test]
    public void GitIssue271InvalidDimension()
    {
        using var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_271_InvalidDimension.xlsx"));
        var dataSet = excelReader.AsDataSet();
        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(3));
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(9));
    }

    [Test]
    public void GitIssue289CompoundDocumentEncryptedWithDefaultPassword()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue289.xlsx"));
        reader.Read();
        Assert.That(reader.GetValue(0), Is.EqualTo("aaaaaaa"));
    }

    [Test]
    public void GitIssue301IgnoreCase()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_301_IgnoreCase.xlsx"));
        DataTable result = reader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(10));
        Assert.That(result.Columns.Count, Is.EqualTo(10));
        Assert.That(result.Rows[1][0], Is.EqualTo("10x10"));
        Assert.That(result.Rows[9][9], Is.EqualTo("10x27"));
    }

    [Test]
    public void GitIssue319InlineRichText()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue319.xlsx"));
        var result = reader.AsDataSet().Tables[0];

        Assert.That(result.Rows[0][0], Is.EqualTo("Text1"));
    }

    [Test]
    public void GitIssue324MultipleRowElementsPerRow()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_324.xlsx"));
        var result = reader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(20));
        Assert.That(result.Columns.Count, Is.EqualTo(13));

        Assert.That(result.Rows[10].ItemArray, Is.EqualTo(new object[] { DBNull.Value, DBNull.Value, "Other", 191036.15, 194489.45, 66106.32, 37167.88, 102589.54, 57467.94, 130721.93, 150752.67, 76300.69, 67024.6 }));
    }

    [Test]
    public void GitIssue354()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("test_git_issue_354.xlsx"));
        var result = reader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(1));
        Assert.That(result.Rows[0][0], Is.EqualTo("cell data"));
    }

    [Test]
    public void GitIssue385Backslash()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_385_backslash.xlsx"));
        var result = reader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(10));
        Assert.That(result.Columns.Count, Is.EqualTo(10));
        Assert.That(result.Rows[1][0], Is.EqualTo("10x10"));
        Assert.That(result.Rows[9][9], Is.EqualTo("10x27"));
    }

    /// <summary>
    /// This test is to ensure that we get the same results from an xls saved in excel vs open office.
    /// </summary>
    [Test]
    public void TestOpenOfficeSavedInExcel()
    {
        using IExcelDataReader excelReader = OpenReader("Test_Excel_OpenOffice");
        AssertUtilities.DoOpenOfficeTest(excelReader);
    }

    [Test]
    public void GitIssue454HandleDuplicateNumberFormats()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue454.xlsx"));
        reader.Read();
    }

    [Test]
    public void GitIssue486TransformValue()
    {
        using var reader = OpenReader("Test_git_issue_486");
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            // If you set data type to an enum and import DataSet transforms the boxed enum values to boxed ints instead
            UseColumnDataType = false,
            ConfigureDataTable = _ => new ExcelDataTableConfiguration()
            {
                UseHeaderRow = true,
                TransformValue = (transformReader, n, value) =>
                {
                    var error = transformReader.GetCellError(n);
                    if (error != null)
                    {
                        return error;
                    }

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
    public void GitIssue608()
    {
        using var reader = OpenReader("Test_git_issue_608");
        var dataSet = reader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[0].ItemArray, Is.EqualTo(new[] { "Column1", "Column2", "Column3" }));
    }

    [Test]
    public void GitIssue629()
    {
        using var reader = OpenReader("Test_git_issue_629");
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].Rows[9].ItemArray[0], Is.EqualTo("Transref. AM"));
    }

    [Test]
    public void GitIssue425()
    {
        using var reader = OpenReader("Test_git_issue_425");
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].Rows[0].ItemArray[0], Is.EqualTo("   text    "));
        Assert.That(dataSet.Tables[0].Rows[1].ItemArray[0], Is.EqualTo("text"));
        Assert.That(dataSet.Tables[0].Rows[2].ItemArray[0], Is.EqualTo("text    text"));
    }

    [Test]
    public void GitIssue649()
    {
        using var reader = OpenReader("Test_git_issue_649_Date1904_Parsings");
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].Rows[11].ItemArray[12], Is.EqualTo(new DateTime(2023, 9, 1)));
    }

    [Test]
    public void GitIssue518MultipleHeaderRows()
    {
        using (var reader = OpenReader("Test_git_issue_518"))
        {
            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true,
                    ReadHeader = self =>
                    {
                        List<string> headerNames = [];

                        // read first header row
                        for (var i = 0; i < self.FieldCount; i++)
                        {
                            var s = Convert.ToString(self.GetValue(i));
                            headerNames.Add(s);
                        }

                        // append second header row
                        if (!self.Read()) 
                        {
                            throw new Exception();
                        }

                        var result = new Dictionary<int, string>(self.FieldCount);
                        for (var i = 0; i < self.FieldCount; i++)
                        {
                            var first = headerNames[i];
                            var second = Convert.ToString(self.GetValue(i));
                            string name;
                            if (first.Length == 0) 
                            {
                                name = second;
                            } 
                            else if (second.Length == 0) 
                            {
                                name = first;
                            } 
                            else 
                            {
                                name = first + " " + second;
                            }

                            if (string.IsNullOrEmpty(name))
                            {
                                name = "Column" + i;
                            }

                            result.Add(i, name);
                        }

                        return result;
                    }
                }
            });

            var columns = dataSet.Tables[0].Columns;
            Assert.That(columns[0].ColumnName.ToString(), Is.EqualTo("ColName1 A"));
            Assert.That(columns[1].ColumnName.ToString(), Is.EqualTo("ColName1 B"));
            Assert.That(columns[2].ColumnName.ToString(), Is.EqualTo("ColName2 B"));
            Assert.That(columns[3].ColumnName.ToString(), Is.EqualTo("FirstOnly"));
            Assert.That(columns[4].ColumnName.ToString(), Is.EqualTo("SecondOnly"));
            Assert.That(columns[5].ColumnName.ToString(), Is.EqualTo("Another One"));
        }
    }

    [Test]
    public void GitIssue642_ActiveSheet()
    {
        using var reader = OpenReader("Test_git_issue_642");
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            FilterSheet = (tableReader, sheetIndex) => tableReader.IsActiveSheet
        });
        Assert.That(reader.ActiveSheet, Is.EqualTo(5));
        Assert.That(dataSet.Tables[0].TableName, Is.EqualTo("List6"));
    }

    [Test]
    public void GitIssue642_ActiveSheet_SingleWorksheet()
    {
        using var reader = OpenReader("Test_git_issue_642onesheet");
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            FilterSheet = (tableReader, sheetIndex) => tableReader.IsActiveSheet
        });
        Assert.That(reader.ActiveSheet, Is.EqualTo(0));
        Assert.That(dataSet.Tables[0].TableName, Is.EqualTo("List1"));
    }

    [Test]
    public void GitIssue700_AlignmentEnumParsing()
    {
        using var reader = OpenReader("Test_git_issue_700_CellAlignments");
        reader.Read();

        var general = reader.GetCellStyle(0);
        var left = reader.GetCellStyle(1);
        var center = reader.GetCellStyle(2);
        var right = reader.GetCellStyle(3);

        Assert.That(general.HorizontalAlignment, Is.EqualTo(HorizontalAlignment.General));
        Assert.That(left.HorizontalAlignment, Is.EqualTo(HorizontalAlignment.Left));
        Assert.That(center.HorizontalAlignment, Is.EqualTo(HorizontalAlignment.Center));
        Assert.That(right.HorizontalAlignment, Is.EqualTo(HorizontalAlignment.Right));
    }

    [TestCase("Test_git_issue_711_OpenXml_Row_height_parsing")]
    [TestCase("Test_git_issue_711_OpenXml_Row_height_parsing_negative_default_height")]
    public void GitIssue711_RowHeightParsing(string filename)
    {
        using var reader = OpenReader(filename);
        var expectedRowHeights = new List<double>
        {
            15, // 0.  -
            15, // 1.  hidden="0"
            0,  // 2.  hidden="1"
            0,  // 3.  ht="0"
            0,  // 4.  ht="-0"
            0,  // 5.  ht="0" hidden="0"
            0,  // 6.  ht="-0" hidden="0"
            0,  // 7.  ht="0" hidden="1"
            0,  // 8.  ht="-0" hidden="1"
            20, // 9.  ht="20"
            20, // 10. ht="-20"
            20, // 11. ht="20" hidden="0"
            20, // 12. ht="-20" hidden="0"
            0,  // 13. ht="20" hidden="1"
            0,  // 14. ht="-20" hidden="1"
            15, // 15. ht="string"
            15, // 16. ht="string" hidden="0"
            0,  // 17. ht="string" hidden="1"
        };
        var actualRowsHeights = new List<double>();
        while (reader.Read())
        {
            actualRowsHeights.Add(reader.RowHeight);
        }

        Assert.That(actualRowsHeights, Is.EqualTo(expectedRowHeights));
    }

    protected override IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null) 
    => ExcelReaderFactory.CreateOpenXmlReader(stream, configuration);

    protected override Stream OpenStream(string name) => Configuration.GetTestWorkbook(name + ".xlsx");
}
