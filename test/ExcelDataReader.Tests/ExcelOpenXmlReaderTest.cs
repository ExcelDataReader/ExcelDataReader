using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using NUnit.Framework;

namespace ExcelDataReader.Tests
{

    public class ExcelOpenXmlReaderTest : ExcelOpenXmlReaderBase
    {
        protected override IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null)
        {
            return ExcelReaderFactory.CreateOpenXmlReader(stream, configuration);
        }

        protected override Stream OpenStream(string name)
        {
            return Configuration.GetTestWorkbook(name + ".xlsx");
        }

        /// <inheritdoc />
        protected override DateTime GitIssue82TodayDate => new DateTime(2013, 4, 19);

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

            Assert.AreEqual("Invalid file signature.", exception.Message);
        }

        [Test]
        public void Issue4145()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_4145.xlsx")))
            {
                Assert.DoesNotThrow(() => excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration));

                while (excelReader.Read())
                {
                }
            }
        }

        [Test]
        public void IssueFileLock5161()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestMultiSheet.xlsx")))
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

        [Test]
        public void Issue11522OpenXml()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11522_OpenXml.xlsx")))
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_googlesourced.xlsx")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual("9583638582", dataSet.Tables[0].Rows[0][0].ToString());
                Assert.AreEqual(4, dataSet.Tables[0].Rows.Count);
                Assert.AreEqual(6, dataSet.Tables[0].Columns.Count);
            }
        }

        [Test]
        public void TestIssue12667GoogleExportMissingColumns()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_12667_GoogleExport_MissingColumns.xlsx")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(7, dataSet.Tables[0].Columns.Count); // 6 with data + 1 that is present but no data in it
                Assert.AreEqual(0, dataSet.Tables[0].Rows.Count);
            }
        }

        [Test]
        public void IssueGit142()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_Issue_142.xlsx")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual(4, dataSet.Tables[0].Columns.Count);
            }
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
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_NoStyles_NoRAttribute.xlsx")))
            {
                DataSet result = excelReader.AsDataSet();

                Assert.IsTrue(result.Tables.Count > 0);
                Assert.AreEqual(39, result.Tables[0].Rows.Count);
                Assert.AreEqual(18, result.Tables[0].Columns.Count);
                Assert.AreEqual("ROW NUMBER 5", result.Tables[0].Rows[4][4].ToString());

                excelReader.Close();
            }
        }

        [Test]
        public void NoDimensionOrCellReferenceAttribute()
        {
            // 20170306_Daily Package GPR 250 Index EUR Overview.xlsx
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("NoDimensionOrCellReferenceAttribute.xlsx")))
            {
                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(2, result.Tables.Count);
                Assert.AreEqual(8, result.Tables[0].Columns.Count, "Sheet0 Columns");
                Assert.AreEqual(7, result.Tables[0].Rows.Count, "Sheet0 Rows");

                Assert.AreEqual(8, result.Tables[1].Columns.Count, "Sheet1 Columns");
                Assert.AreEqual(20, result.Tables[1].Rows.Count, "Sheet1 Rows");
            }
        }

        [Test]
        public void CellValueIso8601Date()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_221.xlsx")))
            {
                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(new DateTime(2017, 3, 16), result.Tables[0].Rows[0][0]);
            }
        }

        [Test]
        public void CellFormat49()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Format49_@.xlsx")))
            {
                DataSet result = excelReader.AsDataSet();

                // ExcelDataReader used to convert numbers formatted with NumFmtId=49/@ to culture-specific strings.
                // This behaviour changed in v3 to return the original value:
                // Assert.That(result.Tables[0].Rows[0].ItemArray, Is.EqualTo(new[] { "2010-05-05", "1.1", "2,2", "123", "2,2" }));
                Assert.That(result.Tables[0].Rows[0].ItemArray, Is.EqualTo(new object[] { "2010-05-05", "1.1", 2.2000000000000002D, 123.0D, "2,2" }));
            }
        }

        [Test]
        public void GitIssue97()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("fillreport.xlsx")))
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

        [Test]
        public void GitIssue68NullSheetPath()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_68_NullSheetPath.xlsm")))
            {
                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(2, result.Tables[0].Columns.Count);
                Assert.AreEqual(1, result.Tables[0].Rows.Count);

            }
        }

        [Test]
        public void GitIssue53CachedFormulaStringType()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_53_Cached_Formula_String_Type.xlsx")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                // Ensure that parseable, numeric cached formula values are read as a double
                Assert.IsInstanceOf<double>(dataSet.Tables[0].Rows[0][2]);
                Assert.AreEqual(3D, dataSet.Tables[0].Rows[0][2]);

                // Ensure that non-parseable, non-numeric cached formula values are read as a string
                Assert.IsInstanceOf<string>(dataSet.Tables[0].Rows[1][2]);
                Assert.AreEqual("AB", dataSet.Tables[0].Rows[1][2]);

                // Ensure that parseable, non-numeric cached formula values are read as a string
                Assert.IsInstanceOf<string>(dataSet.Tables[0].Rows[2][2]);
                Assert.AreEqual("1,", dataSet.Tables[0].Rows[2][2]);
            }
        }


        [Test]
        public void GitIssue271InvalidDimension()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_271_InvalidDimension.xlsx")))
            {
                var dataSet = excelReader.AsDataSet();
                Assert.AreEqual(3, dataSet.Tables[0].Columns.Count);
                Assert.AreEqual(9, dataSet.Tables[0].Rows.Count);
            }
        }

        [Test]
        public void GitIssue289CompoundDocumentEncryptedWithDefaultPassword()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue289.xlsx")))
            {
                reader.Read();
                Assert.AreEqual("aaaaaaa", reader.GetValue(0));
            }
        }

        [Test]
        public void GitIssue301IgnoreCase()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_301_IgnoreCase.xlsx")))
            {
                DataTable result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(10, result.Rows.Count);
                Assert.AreEqual(10, result.Columns.Count);
                Assert.AreEqual("10x10", result.Rows[1][0]);
                Assert.AreEqual("10x27", result.Rows[9][9]);
            }
        }

        [Test]
        public void GitIssue319InlineRichText()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue319.xlsx")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual("Text1", result.Rows[0][0]);
            }
        }

        [Test]
        public void GitIssue324MultipleRowElementsPerRow()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_324.xlsx")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(20, result.Rows.Count);
                Assert.AreEqual(13, result.Columns.Count);

                Assert.That(result.Rows[10].ItemArray, Is.EqualTo(new object[] { DBNull.Value, DBNull.Value, "Other", 191036.15, 194489.45, 66106.32, 37167.88, 102589.54, 57467.94, 130721.93, 150752.67, 76300.69, 67024.6 }));
            }
        }

        [Test]
        public void GitIssue354()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("test_git_issue_354.xlsx")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(1, result.Rows.Count);
                Assert.AreEqual("cell data", result.Rows[0][0]);
            }
        }

        [Test]
        public void GitIssue385Backslash()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_385_backslash.xlsx")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(10, result.Rows.Count);
                Assert.AreEqual(10, result.Columns.Count);
                Assert.AreEqual("10x10", result.Rows[1][0]);
                Assert.AreEqual("10x27", result.Rows[9][9]);
            }
        }

        /// <summary>
        /// This test is to ensure that we get the same results from an xls saved in excel vs open office
        /// </summary>
        [Test]
        public void TestOpenOfficeSavedInExcel()
        {
            using (IExcelDataReader excelReader = OpenReader("Test_Excel_OpenOffice"))
            {
                AssertUtilities.DoOpenOfficeTest(excelReader);
            }
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
            using (var reader = OpenReader("Test_git_issue_486"))
            {
                var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
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

                Assert.AreEqual("REF", dataSet.Tables[0].Rows[0][0].ToString());
                Assert.AreEqual("REF", dataSet.Tables[0].Rows[0][1].ToString());

                Assert.AreEqual("NAME", dataSet.Tables[0].Rows[1][0].ToString());
                Assert.AreEqual("NAME", dataSet.Tables[0].Rows[1][1].ToString());
            }
        }        
    }
}
