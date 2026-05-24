using System.Data;

namespace ExcelDataReader.Tests;

public class ExcelOpenXmlReaderTest : ExcelOpenXmlReaderBase
{
    /// <inheritdoc />
    protected override DateTime Issue82_TodayDate => new(2013, 4, 19);

    /// <summary>
    /// Regression test for https://github.com/ExcelDataReader/ExcelDataReader/issues/741.
    /// The static XmlNameTable in ZipWorker is shared across all XmlReader instances.
    /// NameTable is not thread-safe: concurrent Add() calls from multiple threads corrupt
    /// its internal hash table, causing GetAttribute() to return null and subsequently
    /// uint.Parse(null) to throw ArgumentNullException.
    /// </summary>
    [Test]
    public void Issue741_ConcurrentOpenXml()
    {
        const int threadCount = 8;
        var barrier = new System.Threading.Barrier(threadCount);
        var exceptions = new System.Collections.Concurrent.ConcurrentBag<Exception>();

        var tasks = Enumerable.Range(0, threadCount).Select(i => System.Threading.Tasks.Task.Run(() =>
        {
            try
            {
                // Synchronise all threads to hit CreateReader at the same instant,
                // maximising the chance of triggering the NameTable race condition.
                barrier.SignalAndWait();

                using var stream = Configuration.GetTestWorkbook("10x10.xlsx");
                using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                while (reader.Read())
                {
                    for (int col = 0; col < reader.FieldCount; col++)
                        reader.GetValue(col);
                }
            }
            catch (Exception ex)
            {
                exceptions.Add(ex);
            }
        })).ToArray();

        System.Threading.Tasks.Task.WaitAll(tasks);

        Assert.That(exceptions, Is.Empty, $"Exception(s) thrown during concurrent reads:\n{string.Join("\n", exceptions.Select(e => e.ToString()))}");
    }

    [Test]
    public void FailTest()
    {
        var expectedException = typeof(Exceptions.HeaderException);

        var exception = Assert.Throws(expectedException, () =>
            {
                using (ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("FailBinary.xls")))
                {
                }
            });

        Assert.That(exception.Message, Is.EqualTo("Invalid file signature."));
    }

    [Test]
    public void Issue4145()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("OldIssue4145.xlsx"));
        Assert.DoesNotThrow(() => excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration));

        while (excelReader.Read())
        {
        }
    }

    [Test]
    public void IssueFileLock5161()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("MultiSheet.xlsx"));
        
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
    public void Issue11522_OpenXml()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("OldIssue11522_OpenXml.xlsx"));
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
                zipper.Extract(Configuration.GetTestWorkbook("Chess"));
                Assert.AreEqual(false, Directory.Exists(zipper.TempPath));
                Assert.AreEqual(false, zipper.IsValid);

                //this one is valid so we expect to find the files
                zipper.Extract(Configuration.GetTestWorkbook("OpenXml"));

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
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Googlesourced.xlsx"));
        var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

        Assert.That(dataSet.Tables[0].Rows[0][0].ToString(), Is.EqualTo("9583638582"));
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(4));
        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(6));
    }

    [Test]
    public void TestIssue12667_GoogleExportMissingColumns()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("OldIssue12667_GoogleExportMissingColumns.xlsx"));
        var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(6)); // 6 with data + 1 that is present but no data in it
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(0));
    }

    [Test]
    public void IssueGit142()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Issue142.xlsx"));
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
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("NoStylesNoRAttribute.xlsx"));
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
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Issue221.xlsx"));
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
        using var reader = OpenReader("FullWidthSpace");
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].Rows[0].ItemArray[0], Is.EqualTo("\u3000\u3000text\u3000\u3000"));
        Assert.That(dataSet.Tables[0].Rows[1].ItemArray[0], Is.EqualTo("text\u3000"));
        Assert.That(dataSet.Tables[0].Rows[2].ItemArray[0], Is.EqualTo("\u3000text"));
    }
 
    [Test]
    public void Issue97()
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
    public void Issue68_NullSheetPath()
    {
        using var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Issue68_NullSheetPath.xlsm"));
        DataSet result = excelReader.AsDataSet();
        Assert.That(result.Tables[0].Columns.Count, Is.EqualTo(2));
        Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(1));
    }

    [Test]
    public void Issue53_CachedFormulaStringType()
    {
        using var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Issue53_CachedFormulaStringType.xlsx"));
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
    public void Issue271_InvalidDimension()
    {
        using var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Issue271_InvalidDimension.xlsx"));
        var dataSet = excelReader.AsDataSet();
        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(3));
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(9));
    }

    [Test]
    public void Issue289_CompoundDocumentEncryptedWithDefaultPassword()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Issue289.xlsx"));
        reader.Read();
        Assert.That(reader.GetValue(0), Is.EqualTo("aaaaaaa"));
    }

    [Test]
    public void Issue301_IgnoreCase()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Issue301_IgnoreCase.xlsx"));
        DataTable result = reader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(10));
        Assert.That(result.Columns.Count, Is.EqualTo(10));
        Assert.That(result.Rows[1][0], Is.EqualTo("10x10"));
        Assert.That(result.Rows[9][9], Is.EqualTo("10x27"));
    }

    [Test]
    public void Issue319_InlineRichText()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Issue319.xlsx"));
        var result = reader.AsDataSet().Tables[0];

        Assert.That(result.Rows[0][0], Is.EqualTo("Text1"));
    }

    [Test]
    public void Issue324_MultipleRowElementsPerRow()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Issue324.xlsx"));
        var result = reader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(20));
        Assert.That(result.Columns.Count, Is.EqualTo(13));

        Assert.That(result.Rows[10].ItemArray, Is.EqualTo(new object[] { DBNull.Value, DBNull.Value, "Other", 191036.15, 194489.45, 66106.32, 37167.88, 102589.54, 57467.94, 130721.93, 150752.67, 76300.69, 67024.6 }));
    }

    [Test]
    public void Issue354()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Issue354.xlsx"));
        var result = reader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(1));
        Assert.That(result.Rows[0][0], Is.EqualTo("cell data"));
    }

    [Test]
    public void Issue385_Backslash()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Issue385_Backslash.xlsx"));
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
    public void OpenOfficeSavedInExcel()
    {
        using IExcelDataReader excelReader = OpenReader("ExcelOpenOffice");
        AssertUtilities.DoOpenOfficeTest(excelReader);
    }

    [Test]
    public void Issue454_HandleDuplicateNumberFormats()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Issue454.xlsx"));
        reader.Read();
    }

    [Test]
    public void Issue486_TransformValue()
    {
        using var reader = OpenReader("Issue486");
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
    public void Issue608()
    {
        using var reader = OpenReader("Issue608");
        var dataSet = reader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[0].ItemArray, Is.EqualTo(new[] { "Column1", "Column2", "Column3" }));
    }

    [Test]
    public void Issue629()
    {
        using var reader = OpenReader("Issue629");
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].Rows[9].ItemArray[0], Is.EqualTo("Transref. AM"));
    }

    [Test]
    public void Issue425()
    {
        using var reader = OpenReader("Issue425");
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].Rows[0].ItemArray[0], Is.EqualTo("   text    "));
        Assert.That(dataSet.Tables[0].Rows[1].ItemArray[0], Is.EqualTo("text"));
        Assert.That(dataSet.Tables[0].Rows[2].ItemArray[0], Is.EqualTo("text    text"));
    }

    [Test]
    public void Issue649()
    {
        using var reader = OpenReader("Issue649_Date1904Parsings");
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].Rows[11].ItemArray[12], Is.EqualTo(new DateTime(2023, 9, 1)));
    }

    [Test]
    public void Issue518_MultipleHeaderRows()
    {
        using (var reader = OpenReader("Issue518"))
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
    public void Issue642_ActiveSheet()
    {
        using var reader = OpenReader("Issue642");
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            FilterSheet = (tableReader, sheetIndex) => tableReader.IsActiveSheet
        });
        Assert.That(reader.ActiveSheet, Is.EqualTo(5));
        Assert.That(dataSet.Tables[0].TableName, Is.EqualTo("List6"));
    }

    [Test]
    public void Issue642_ActiveSheet_SingleWorksheet()
    {
        using var reader = OpenReader("Issue642_Onesheet");
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            FilterSheet = (tableReader, sheetIndex) => tableReader.IsActiveSheet
        });
        Assert.That(reader.ActiveSheet, Is.EqualTo(0));
        Assert.That(dataSet.Tables[0].TableName, Is.EqualTo("List1"));
    }

    [Test]
    public void Issue700_AlignmentEnumParsing()
    {
        using var reader = OpenReader("Issue700_CellAlignments");
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

    [TestCase("Issue711_OpenXmlRowHeightParsing")] // defaultHeight="15"
    [TestCase("Issue711_OpenXmlRowHeightParsingNegativeDefaultHeight")] // defaultHeight="-15"
    public void Issue711_RowHeightParsing(string filename)
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

    [Test]
    public void Issue461_Format14WithEnUsCultureReturnsCorrectFormatString()
    {
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetNumberFormatString(0, new System.Globalization.CultureInfo("en-US")), Is.EqualTo("m/d/yyyy"));
    }

    [Test]
    public void Issue461_Format14WithEnGbCultureReturnsCorrectFormatString()
    {
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetNumberFormatString(0, new System.Globalization.CultureInfo("en-GB")), Is.EqualTo("dd/mm/yyyy"));
    }

    [Test]
    public void Issue461_Format14DefaultBehaviorUnchanged()
    {
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetNumberFormatString(0), Is.EqualTo("d/m/yyyy"));
    }

    [Test]
    public void Issue461_Formats15To17WithEnUsCultureUsesSlashSeparator()
    {
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
        var enUs = new System.Globalization.CultureInfo("en-US");

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetNumberFormatString(1, enUs), Is.EqualTo("d/mmm/yy"));  // format 15
        Assert.That(reader.GetNumberFormatString(2, enUs), Is.EqualTo("d/mmm"));     // format 16
        Assert.That(reader.GetNumberFormatString(3, enUs), Is.EqualTo("mmm/yy"));    // format 17
    }

    [Test]
    public void Issue461_Formats15To17WithDeDeCultureUsesDotSeparator()
    {
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
        var deDE = new System.Globalization.CultureInfo("de-DE");

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetNumberFormatString(1, deDE), Is.EqualTo("d.mmm.yy"));  // format 15
        Assert.That(reader.GetNumberFormatString(2, deDE), Is.EqualTo("d.mmm"));     // format 16
        Assert.That(reader.GetNumberFormatString(3, deDE), Is.EqualTo("mmm.yy"));    // format 17
    }

    [Test]
    public void Issue461_Formats15To17DefaultBehaviorUnchanged()
    {
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetNumberFormatString(1), Is.EqualTo("d-mmm-yy"));  // format 15
        Assert.That(reader.GetNumberFormatString(2), Is.EqualTo("d-mmm"));     // format 16
        Assert.That(reader.GetNumberFormatString(3), Is.EqualTo("mmm-yy"));    // format 17
    }

    [Test]
    public void Issue461_CellValueIsDateTimeRegardlessOfCulture()
    {
        // Whether a provider is passed or not, cells with format 14 should be returned as DateTime
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetValue(0), Is.EqualTo(new DateTime(2023, 1, 1)));
    }

    [Test]
    public void Issue461_Format20WithEnUsCultureFallsBackToSpecString()
    {
        // en-US uses 12h AM/PM time, so index 20 (explicitly non-AM/PM) falls back to the spec string
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetNumberFormatString(4, new System.Globalization.CultureInfo("en-US")), Is.EqualTo("h:mm"));
    }

    [Test]
    public void Issue461_Format20WithDeDeCultureReturnsTwentyFourHourTime()
    {
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetNumberFormatString(4, new System.Globalization.CultureInfo("de-DE")), Is.EqualTo("hh:mm"));
    }

    [Test]
    public void Issue461_Format20DefaultBehaviorUnchanged()
    {
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetNumberFormatString(4), Is.EqualTo("h:mm"));
    }

    [Test]
    public void Issue461_Format21WithEnUsCultureFallsBackToSpecString()
    {
        // en-US uses 12h AM/PM time, so index 21 (explicitly non-AM/PM) falls back to the spec string
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetNumberFormatString(5, new System.Globalization.CultureInfo("en-US")), Is.EqualTo("h:mm:ss"));
    }

    [Test]
    public void Issue461_Format21WithDeDeCultureReturnsTwentyFourHourTime()
    {
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetNumberFormatString(5, new System.Globalization.CultureInfo("de-DE")), Is.EqualTo("hh:mm:ss"));
    }

    [Test]
    public void Issue461_Format21DefaultBehaviorUnchanged()
    {
        using var stream = Configuration.GetTestWorkbook("Issue461.xlsx");
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetNumberFormatString(5), Is.EqualTo("h:mm:ss"));
    }

    protected override IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null) 
    => ExcelReaderFactory.CreateOpenXmlReader(stream, configuration);

    protected override Stream OpenStream(string name) => Configuration.GetTestWorkbook(name + ".xlsx");
}
