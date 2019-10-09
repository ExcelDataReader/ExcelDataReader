using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

using NUnit.Framework;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestCleanup = NUnit.Framework.TearDownAttribute;
using TestInitialize = NUnit.Framework.SetUpAttribute;
using TestMethod = NUnit.Framework.TestAttribute;

namespace ExcelDataReader.Tests
{
    [TestClass]
    public class ExcelOpenXmlReaderTest : ExcelTestBase
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

        [TestMethod]
        public void Issue11516WorkbookWithSingleSheetShouldNotReturnEmptyDataset()
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11516_Single_Tab.xlsx")))
            {
                Assert.AreEqual(1, reader.ResultsCount);

                DataSet dataSet = reader.AsDataSet();

                Assert.IsTrue(dataSet != null);
                Assert.AreEqual(1, dataSet.Tables.Count);
                Assert.AreEqual(260, dataSet.Tables[0].Rows.Count);
                Assert.AreEqual(29, dataSet.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void AsDataSetTestReadSheetNames()
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("TestOpenXml.xlsx")))
            {
                Assert.AreEqual(3, reader.ResultsCount);

                DataSet dataSet = reader.AsDataSet();

                Assert.IsTrue(dataSet != null);
                Assert.AreEqual(3, dataSet.Tables.Count);
                Assert.AreEqual(7, dataSet.Tables["Sheet1"].Rows.Count);
                Assert.AreEqual(11, dataSet.Tables["Sheet1"].Columns.Count);
            }
        }

        [TestMethod]
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

        [TestMethod]
        public void IssueDateAndTime1468Test()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Encoding_Formula_Date_1520.xlsx")))
            {
                DataSet dataSet = excelReader.AsDataSet();

                string val1 = new DateTime(2009, 05, 01).ToShortDateString();
                string val2 = DateTime.Parse(dataSet.Tables[0].Rows[1][1].ToString()).ToShortDateString();

                Assert.AreEqual(val1, val2);

                val1 = new DateTime(2009, 1, 1, 11, 0, 0).ToShortTimeString();
                val2 = DateTime.Parse(dataSet.Tables[0].Rows[2][4].ToString()).ToShortTimeString();

                Assert.AreEqual(val1, val2);
            }
        }

        [TestMethod]
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

        [TestMethod]
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

        [TestMethod]
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
                [TestMethod]
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

        [TestMethod]
        public void Issue11773Exponential()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11773_Exponential.xlsx")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(2566.37168141593D, double.Parse(dataSet.Tables[0].Rows[0][6].ToString()));
            }
        }

        [TestMethod]
        public void Issue11773ExponentialCommas()
        {
#if NETCOREAPP1_0
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");
#else
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE", false);
#endif

            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_11773_Exponential.xlsx")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(2566.37168141593D, double.Parse(dataSet.Tables[0].Rows[0][6].ToString()));
            }
        }

        [TestMethod]
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

        [TestMethod]
        public void TestIssue12667GoogleExportMissingColumns()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_12667_GoogleExport_MissingColumns.xlsx")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual(7, dataSet.Tables[0].Columns.Count); // 6 with data + 1 that is present but no data in it
                Assert.AreEqual(0, dataSet.Tables[0].Rows.Count);
            }
        }

        /// <summary>
        /// Makes sure that we can read data from the first row of last sheet
        /// </summary>
        [TestMethod]
        public void Issue12271NextResultSet()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_LotsOfSheets.xlsx")))
            {
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
        }

        [TestMethod]
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
        [TestMethod]
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

        [TestMethod]
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

        [TestMethod]
        public void CellValueIso8601Date()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_221.xlsx")))
            {
                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(new DateTime(2017, 3, 16), result.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
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

        [TestMethod]
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

        [TestMethod]
        public void GitIssue68NullSheetPath()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_68_NullSheetPath.xlsm")))
            {
                DataSet result = excelReader.AsDataSet();
                Assert.AreEqual(2, result.Tables[0].Columns.Count);
                Assert.AreEqual(1, result.Tables[0].Rows.Count);

            }
        }

        [TestMethod]
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

        [TestMethod]
        public void GitIssue14InvalidOADate()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_14_InvalidOADate.xlsx")))
            {
                var dataSet = excelReader.AsDataSet();

                // Test out of range double formatted as date returns double
                Assert.AreEqual(1000000000000D, dataSet.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void GitIssue241FirstOddEven()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_224_firstoddeven.xlsx")))
            {
                Assert.That(reader.HeaderFooter, Is.Not.Null);

                Assert.That(reader.HeaderFooter?.HasDifferentFirst, Is.True, "HasDifferentFirst");
                Assert.That(reader.HeaderFooter?.HasDifferentOddEven, Is.True, "HasDifferentOddEven");

                Assert.That(reader.HeaderFooter?.FirstHeader, Is.EqualTo("&CFirst header center"), "First Header");
                Assert.That(reader.HeaderFooter?.FirstFooter, Is.EqualTo("&CFirst footer center"), "First Footer");
                Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft едц &T&COdd page header&RRight  едц &P"), "Odd Header");
                Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft едц &P&COdd Footer едц &P&RRight едц &D"), "Odd Footer");
                Assert.That(reader.HeaderFooter?.EvenHeader, Is.EqualTo("&L&A&CEven page header"), "Even Header");
                Assert.That(reader.HeaderFooter?.EvenFooter, Is.EqualTo("&CEven page footer"), "Even Footer");
            }
        }

        [TestMethod]
        public void GitIssue242StandardEncryption()
        {
            // OpenXml standard encryption aes128+sha1
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("standard_AES128_SHA1_ECB_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml standard encryption aes192+sha1
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("standard_AES192_SHA1_ECB_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml standard encryption aes256+sha1
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("standard_AES256_SHA1_ECB_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }
        }

        [TestMethod]
        public void GitIssue242AgileEncryption()
        {
            // OpenXml agile encryption aes128+md5+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES128_MD5_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes128+sha1+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES128_SHA1_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes128+sha384+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES128_SHA384_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes128+sha512+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES128_SHA512_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes192+sha512+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES192_SHA512_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption aes256+sha512+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_AES256_SHA512_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption 3des+sha384+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_DESede_SHA384_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // The following encryptions do not exist on netstandard1.3
#if NET20 || NET45 || NETCOREAPP2_0
            // OpenXml agile encryption des+md5+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_DES_MD5_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // OpenXml agile encryption rc2+sha1+cbc
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                Configuration.GetTestWorkbook("agile_RC2_SHA1_CBC_pwd_password.xlsx"),
                new ExcelReaderConfiguration() { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }
#endif
        }

        [TestMethod]
        public void OpenXmlThrowsInvalidPassword()
        {
            Assert.Throws<Exceptions.InvalidPasswordException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                    Configuration.GetTestWorkbook("agile_AES128_MD5_CBC_pwd_password.xlsx"),
                    new ExcelReaderConfiguration() { Password = "wrongpassword" }))
                {
                    reader.Read();
                }
            });

            Assert.Throws<Exceptions.InvalidPasswordException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(
                    Configuration.GetTestWorkbook("agile_AES128_MD5_CBC_pwd_password.xlsx")))
                {
                    reader.Read();
                }
            });
        }

        [TestMethod]
        public void OpenXmlThrowsEmptyZipFile()
        {
            Assert.Throws<Exceptions.HeaderException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("EmptyZipFile.xlsx")))
                {
                    reader.Read();
                }
            });
        }

        [TestMethod]
        public void GitIssue265OpenXmlDisposed()
        {
            // Verify the file stream is closed and disposed by the reader
            {
                var stream = Configuration.GetTestWorkbook("Test10x10.xlsx");
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    var _ = excelReader.AsDataSet();
                }

                Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
            }

            // Verify streams used by standard encryption are closed
            {
                var stream = Configuration.GetTestWorkbook("standard_AES128_SHA1_ECB_pwd_password.xlsx");

                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(
                    stream,
                    new ExcelReaderConfiguration() { Password = "password" }))
                {
                    var _ = excelReader.AsDataSet();
                }

                Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
            }

            // Verify streams used by agile encryption are closed
            {
                var stream = Configuration.GetTestWorkbook("agile_AES128_MD5_CBC_pwd_password.xlsx");

                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(
                    stream,
                    new ExcelReaderConfiguration() { Password = "password" }))
                {
                    var _ = excelReader.AsDataSet();
                }

                Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
            }
        }

        [TestMethod]
        public void GitIssue271InvalidDimension()
        {
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_271_InvalidDimension.xlsx")))
            {
                var dataSet = excelReader.AsDataSet();
                Assert.AreEqual(3, dataSet.Tables[0].Columns.Count);
                Assert.AreEqual(9, dataSet.Tables[0].Rows.Count);
            }
        }

        [TestMethod]
        public void GitIssue289CompoundDocumentEncryptedWithDefaultPassword()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue289.xlsx")))
            {
                reader.Read();
                Assert.AreEqual("aaaaaaa", reader.GetValue(0));
            }
        }

        [TestMethod]
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

        [TestMethod]
        public void GitIssue319InlineRichText()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue319.xlsx")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual("Text1", result.Rows[0][0]);
            }
        }

        [TestMethod]
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

        [TestMethod]
        public void GitIssue354()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("test_git_issue_354.xlsx")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(1, result.Rows.Count);
                Assert.AreEqual("cell data", result.Rows[0][0]);
            }
        }

        [TestMethod]
        public void GitIssue364()
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("test_git_issue_364.xlsx")))
            {
                Assert.AreEqual(1, reader.RowCount);
                reader.Read();

                Assert.AreEqual(0, reader.GetNumberFormatIndex(0));
                Assert.AreEqual(-1, reader.GetNumberFormatIndex(1));
                Assert.AreEqual(14, reader.GetNumberFormatIndex(2));
                Assert.AreEqual(164, reader.GetNumberFormatIndex(3));
            }
        }

        [TestMethod]
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
    }
}
