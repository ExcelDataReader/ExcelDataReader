using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader.Exceptions;

using NUnit.Framework;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestCleanup = NUnit.Framework.TearDownAttribute;
using TestInitialize = NUnit.Framework.SetUpAttribute;
using TestMethod = NUnit.Framework.TestAttribute;

namespace ExcelDataReader.Tests
{
    [TestClass]

    public class ExcelBinaryReaderTest : ExcelTestBase
    {
        protected override IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null)
        {
            return ExcelReaderFactory.CreateBinaryReader(stream, configuration);
        }

        protected override Stream OpenStream(string name)
        {
            return Configuration.GetTestWorkbook(name + ".xls");
        }

        /// <inheritdoc />
        protected override DateTime GitIssue82TodayDate => new DateTime(2009, 6, 28);

        [TestMethod]
        public void GitIssue70ExcelBinaryReaderTryConvertOADateTimeFormula()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_70_ExcelBinaryReader_tryConvertOADateTime _convert_dates.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);

                var date = ds.Tables[0].Rows[1].ItemArray[0];

                Assert.AreEqual(new DateTime(2014, 01, 01), date);
            }
        }

        [TestMethod]
        public void GitIssue51ReadCellLabel()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_51.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);

                var value = ds.Tables[0].Rows[0].ItemArray[1];

                Assert.AreEqual("Monetary aggregates (R millions)", value);
            }
        }

        [TestMethod]
        public void GitIssue45()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_45.xls")))
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
        public void Issue1155311570FatIssueOffset()
        {
            void DoTestFatStreamIssue(string sheetId)
            {
                string filePath;
                using (var excelReader1 = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook(sheetId))) // Works.
                {
                    filePath = Configuration.GetTestWorkbookPath(sheetId);
                    Assert.IsNotNull(excelReader1);
                }

                using (var ms1 = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var excelReader2 = ExcelReaderFactory.CreateBinaryReader(ms1)) // Works!
                    Assert.IsNotNull(excelReader2);

                var bytes = File.ReadAllBytes(filePath);
                using (var ms2 = new MemoryStream(bytes))
                using (var excelReader3 = ExcelReaderFactory.CreateBinaryReader(ms2)) // Did not work, but does now
                    Assert.IsNotNull(excelReader3);
            }

            void DoTestFatStreamIssueType2(string sheetId)
            {
                var filePath = Configuration.GetTestWorkbookPath(sheetId);

                using (Stream stream = new MemoryStream(File.ReadAllBytes(filePath)))
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream))
                {
                    // ReSharper disable once AccessToDisposedClosure
                    Assert.DoesNotThrow(() => excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration));
                }
            }

            DoTestFatStreamIssue("Test_Issue_11553_FAT.xls");
            DoTestFatStreamIssueType2("Test_Issue_11570_FAT_1.xls");
            DoTestFatStreamIssueType2("Test_Issue_11570_FAT_2.xls");
        }

        /*[TestMethod]
        public void Test_SSRS()
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_SSRS"));
            DataSet result = excelReader.AsDataSet();
            excelReader.Close();
        }*/

        [TestMethod]
        public void FailTest()
        {
            var exception = Assert.Throws<HeaderException>(() =>
            {
                using (ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestFail_Binary.xls")))
                {
                }
            });

            Assert.AreEqual("Invalid file signature.", exception.Message);
        }

        [TestMethod]
        public void TestOpenOffice()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_OpenOffice.xls")))
            {
                AssertUtilities.DoOpenOfficeTest(excelReader);
            }
        }

        /// <summary>
        /// Issue 11 - OpenOffice files were skipping the first row if IsFirstRowAsColumnNames = false;
        /// </summary>
        [TestMethod]
        public void GitIssue11OpenOfficeRowCount()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_OpenOffice.xls")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.NoColumnNamesConfiguration);
                Assert.AreEqual(34, dataSet.Tables[0].Rows.Count);
            }
        }
        
        [TestMethod]
        public void UncalculatedTest()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Uncalculated.xls")))
            {
                var dataSet = excelReader.AsDataSet();
                Assert.IsNotNull(dataSet);
                Assert.AreNotEqual(dataSet.Tables.Count, 0);
                var table = dataSet.Tables[0];
                Assert.IsNotNull(table);

                Assert.AreEqual("1", table.Rows[1][0].ToString());
                Assert.AreEqual("3", table.Rows[1][2].ToString());
                Assert.AreEqual("3", table.Rows[1][4].ToString());
            }
        }

        [TestMethod]
        public void Issue11570Excel2013()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11570_Excel2013.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual(2, dataSet.Tables[0].Columns.Count);
                Assert.AreEqual(5, dataSet.Tables[0].Rows.Count);

                Assert.AreEqual("1.1.1.2", dataSet.Tables[0].Rows[0][0]);
                Assert.AreEqual(10d, dataSet.Tables[0].Rows[0][1]);

                Assert.AreEqual("1.1.1.15", dataSet.Tables[0].Rows[1][0]);
                Assert.AreEqual(3d, dataSet.Tables[0].Rows[1][1]);

                Assert.AreEqual("2.1.2.23", dataSet.Tables[0].Rows[2][0]);
                Assert.AreEqual(14d, dataSet.Tables[0].Rows[2][1]);

                Assert.AreEqual("2.1.2.31", dataSet.Tables[0].Rows[3][0]);
                Assert.AreEqual(2d, dataSet.Tables[0].Rows[3][1]);

                Assert.AreEqual("2.8.7.30", dataSet.Tables[0].Rows[4][0]);
                Assert.AreEqual(2d, dataSet.Tables[0].Rows[4][1]);
            }
        }

        [TestMethod]
        public void Issue11572CodePage()
        {
            // This test was skipped for a long time as it produced: "System.NotSupportedException : No data is available for encoding 27651."
            // Upon revisiting the underlying cause appears to be fixed
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11572_CodePage.xls")))
            {
                Assert.DoesNotThrow(() => excelReader.AsDataSet());
            }
        }

        /// <summary>
        /// Not fixed yet
        /// </summary>
        [TestMethod]
        public void Issue11545NoIndex()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11545_NoIndex.xls")))
            {
                var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

                Assert.AreEqual("CI2229         ", dataSet.Tables[0].Rows[0][0]);
                Assert.AreEqual("12069E01018A1  ", dataSet.Tables[0].Rows[0][6]);
                Assert.AreEqual(new DateTime(2012, 03, 01), dataSet.Tables[0].Rows[0][8]);
            }
        }

        [TestMethod]
        public void Issue11642ValuesNotLoaded()
        {
            // Excel.Log.Log.InitializeWith<Log4NetLog>();
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11642_ValuesNotLoaded.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual("431113*", dataSet.Tables[2].Rows[29][1].ToString());
                Assert.AreEqual("024807", dataSet.Tables[2].Rows[36][1].ToString());
                Assert.AreEqual("160019", dataSet.Tables[2].Rows[53][1].ToString());
            }
        }

        [TestMethod]
        public void Issue11636BiffStream()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11636_BiffStream.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                // check a couple of values
                Assert.AreEqual("SP011", dataSet.Tables[0].Rows[9][0]);
                Assert.AreEqual(9.9, dataSet.Tables[0].Rows[32][11]);
                Assert.AreEqual(78624.44, dataSet.Tables[1].Rows[27][12]);
            }
        }

        /// <summary>
        /// Not fixed yet
        /// The problem occurs with unseekable stream and logic related to minifat that uses seek
        /// It should probably only use seek if it needs to go backwards, I think at the moment it uses seek all the time
        /// which is probably not good for performance
        /// </summary>
        [TestMethod]
        [Ignore("Not fixed yet")]
        public void Issue1163911644ForwardOnlyStream()
        {
            // Excel.Log.Log.InitializeWith<Log4NetLog>();
            using (var stream = Configuration.GetTestWorkbook("Test_OpenOffice"))
            {
                using (var forwardStream = SeekErrorMemoryStream.CreateFromStream(stream))
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(forwardStream))
                {
                    Assert.DoesNotThrow(() => excelReader.AsDataSet());
                }
            }
        }

        /// <summary>
        /// Not fixed yet
        /// The problem occurs with unseekable stream and logic related to minifat that uses seek
        /// It should probably only use seek if it needs to go backwards, I think at the moment it uses seek all the time
        /// which is probably not good for performance
        /// </summary>
        [TestMethod]
        public void Issue12556Corrupt()
        {
            Assert.Throws<CompoundDocumentException>(() =>
            {
                // Excel.Log.Log.InitializeWith<Log4NetLog>();
                using (var forwardStream = Configuration.GetTestWorkbook("Test_Issue_12556_corrupt.xls"))
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(forwardStream))
                {
                    Assert.DoesNotThrow(() => excelReader.AsDataSet());
                }
            });
        }

        /// <summary>
        /// Some spreadsheets were crashing with index out of range error (from SSRS)
        /// </summary>
        [TestMethod]
        public void TestIssue11818OutOfRange()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11818_OutOfRange.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual("Total Revenue", dataSet.Tables[0].Rows[10][0]);
            }
        }

        [TestMethod]
        public void TestIssue111NoRowRecords()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_111_NoRowRecords.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual(1, dataSet.Tables.Count);
                Assert.AreEqual(12, dataSet.Tables[0].Rows.Count);
                Assert.AreEqual(14, dataSet.Tables[0].Columns.Count);

                Assert.AreEqual(2015.0, dataSet.Tables[0].Rows[7][0]);
            }
        }

        [TestMethod]
        public void TestGitIssue145()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_145.xls")))
            {
                excelReader.Read();
                excelReader.Read();
                excelReader.Read();

                string value = excelReader.GetString(3);

                Assert.AreEqual("Japanese Government Bonds held by the Bank of Japan", value);
            }
        }

        [TestMethod]
        public void TestGitIssue152SheetNameUtf16LeCompressed()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_152.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual("åäöñ", dataSet.Tables[0].TableName);
            }
        }

        [TestMethod]
        public void TestGitIssue152CellUtf16LeCompressed()
        {
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_152.xls")))
            {
                var dataSet = excelReader.AsDataSet();

                Assert.AreEqual("åäöñ", dataSet.Tables[0].Rows[0][0]);
            }
        }

        [TestMethod]
        public void GitIssue158()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_158.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);

                var date = ds.Tables[0].Rows[3].ItemArray[2];

                Assert.AreEqual(new DateTime(2016, 09, 10), date);
            }
        }

        [TestMethod]
        public void GitIssue173()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_173.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);
                Assert.AreEqual(40, ds.Tables.Count);
            }
        }

        [TestMethod]
        public void ReadWriteProtectedStructureUsingStandardEncryption()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("protectedsheet-xxx.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);
                Assert.AreEqual("x", ds.Tables[0].Rows[0][0]);
                Assert.AreEqual(1.4, ds.Tables[0].Rows[1][0]);
            }
        }

        [TestMethod]
        public void TestIncludeTableWithOnlyImage()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestTableOnlyImage_x01oct2016.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.IsNotNull(ds);
                Assert.AreEqual(4, ds.Tables.Count);
            }
        }

        [TestMethod]
        public void AllowFfffAsByteOrder()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_InvalidByteOrderValueInHeader.xls")))
            {
                int tableCount = 0;
                do
                {
                    while (excelReader.Read())
                    {
                    }

                    tableCount++;
                }
                while (excelReader.NextResult());

                Assert.AreEqual(454, tableCount);
            }
        }

        [TestMethod]
        public void HandleRowBlocksWithOutOfOrderCells()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("AllColumnsNotReadInHiddenTable.xls")))
            {
                var ds = excelReader.AsDataSet();

                object[] expected = { "21/09/2015", 1187.5282349881188, 650.8582749049624, 1361.7209439645526, 321.74647548613916, 369.48879457369037 };

                Assert.AreEqual(51, ds.Tables[1].Rows.Count);
                Assert.AreEqual(expected, ds.Tables[1].Rows[1].ItemArray);
            }
        }

        [TestMethod]
        public void HandleRowBlocksWithDifferentNumberOfColumnsAndInvalidDimensions()
        {
            // http://www.ine.cl/canales/chile_estadistico/estadisticas_economicas/edificacion/archivos/xls/edificacion_totalpais_seriehistorica_enero_2017.xls
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("RowWithDifferentNumberOfColumns.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual(256, ds.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void IfNoDimensionDetermineFieldCountByProcessingAllCellColumnIndexes()
        {
            // This xls file has a row record with 256 columns but only values for 6.
            // This test was created when ExcelDataReader incorrectly dropped 8
            // bits off the dimensions' LastColumn in BIFF8 files and relied
            // on scanning to come up with 6 columns. The test was changed to
            // assume valid dimensions:
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_145.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual(5, ds.Tables[0].Columns.Count);
            }
        }

        [TestMethod]
        public void Row1217NotRead()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Row1217NotRead.xls")))
            {
                var ds = excelReader.AsDataSet();
                CollectionAssert.AreEqual(new object[] {
                    DBNull.Value,
                    "Año",
                    "Mes",
                    DBNull.Value, //Merged Cell
                    "Índice",
                    "Variación Mensual",
                    "Variación Acumulada",
                    "Variación en 12 Meses",
                    "Incidencia Mensual",
                    "Incidencia Acumulada", "" +
                    "Incidencia a 12 Meses" }, ds.Tables[0].Rows[1216].ItemArray);
            }
        }

        [TestMethod]
        public void StringContinuationAfterCharacterData()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("StringContinuationAfterCharacterData.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual("商業動態統計速報-平成29年2月分-  統計表", ds.Tables[0].Rows[3][2]);
                Assert.AreEqual("Preliminary Report on the Current Survey of Commerce  ( February,2017 )　Statistics Tables", ds.Tables[0].Rows[4][2]);
                Assert.AreEqual("\nWholesale", ds.Tables[1].Rows[18][9]);
            }
        }

        [TestCase]
        public void Biff3IsSupported()
        {
            using (var stream = Configuration.GetTestWorkbook("biff3.xls"))
            {
                using (var reader = ExcelReaderFactory.CreateBinaryReader(stream))
                {
                    reader.AsDataSet();
                }
            }
        }

        [TestCase]
        public void GitIssue5()
        {
            using (var stream = Configuration.GetTestWorkbook("Test_git_issue_5.xls"))
                Assert.Throws<CompoundDocumentException>(() => ExcelReaderFactory.CreateBinaryReader(stream));
        }

        [TestCase]
        public void Issue2InvalidDimensionRecord()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_2.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual(new[] { "A1", "B1" }, ds.Tables[0].Rows[0].ItemArray);
            }
        }

        [TestCase]
        public void ExcelLibraryNonContinuousMiniStream()
        {
            // Verify the output from the sample code for the ExcelLibrary package parses
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("ExcelLibrary_newdoc.xls")))
            {
                Assert.DoesNotThrow(() => excelReader.AsDataSet());
            }
        }

        [TestCase]
        public void GitIssue184AdditionalFatSectors()
        {
            // Big spreadsheets have additional sectors beyond the header with FAT contents
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("GitIssue_184_FATSectors.xls")))
            {
                DataSet ds = null;
                Assert.DoesNotThrow(() => ds = excelReader.AsDataSet());
                Assert.AreEqual(12, ds.Tables.Count);
                Assert.AreEqual("DATAS (12)", ds.Tables[0].TableName);
                Assert.AreEqual("DATAS (5)", ds.Tables[11].TableName);
            }
        }

        [TestMethod]
        public void RowContentSpreadOverMultipleBlocks()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_217.xls")))
            {
                var ds = excelReader.AsDataSet();
                CollectionAssert.AreEqual(new object[] { "REX GESAMT      ", 484.7929, 142.1032, -0.1656, 5.0315225293000001, 5.0398685515999997, 37.5344725251 }, ds.Tables[2].Rows[10].ItemArray);
            }
        }

        [TestMethod]
        public void GitIssue231NoCodePage()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_231_NoCodePage.xls")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual(11, ds.Tables[0].Columns.Count);
                Assert.AreEqual(5, ds.Tables[0].Rows.Count);
            }
        }

        [TestMethod]
        public void As3XlsBiff2()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF2.xls")))
            {
                DataSet result = excelReader.AsDataSet();
                TestAs3Xls(result);
            }
        }

        [TestMethod]
        public void As3XlsBiff3()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF3.xls")))
            {
                DataSet result = excelReader.AsDataSet();
                TestAs3Xls(result);
            }
        }

        [TestMethod]
        public void As3XlsBiff4()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF4.xls")))
            {
                DataSet result = excelReader.AsDataSet();
                TestAs3Xls(result);
            }
        }

        [TestMethod]
        public void As3XlsBiff5()
        {
            using (var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF5.xls")))
            {
                DataSet result = excelReader.AsDataSet();
                TestAs3Xls(result);
            }
        }

        private static void TestAs3Xls(DataSet result)
        {
            Assert.AreEqual(1, result.Tables[0].Rows[0][0]);
            Assert.AreEqual("Hi", result.Tables[0].Rows[0][1]);
            Assert.AreEqual(10.22D, result.Tables[0].Rows[0][2]);
            Assert.AreEqual(14.754317602356753D, result.Tables[0].Rows[0][3]);
            Assert.AreEqual(21.04107572533686D, result.Tables[0].Rows[0][4]);

            Assert.AreEqual(2, result.Tables[0].Rows[1][0]);
            Assert.AreEqual("How", result.Tables[0].Rows[1][1]);
            Assert.AreEqual(new DateTime(2007, 2, 22), result.Tables[0].Rows[1][2]);

            Assert.AreEqual(3, result.Tables[0].Rows[2][0]);
            Assert.AreEqual("are", result.Tables[0].Rows[2][1]);
            Assert.AreEqual(new DateTime(2002, 1, 19), result.Tables[0].Rows[2][2]);

            Assert.AreEqual("Saturday", result.Tables[0].Rows[3][2]);
            Assert.AreEqual(0.33000000000000002D, result.Tables[0].Rows[4][2]);
            Assert.AreEqual(19, result.Tables[0].Rows[5][2]);
            Assert.AreEqual("Goog", result.Tables[0].Rows[6][2]);
            Assert.AreEqual(12.19D, result.Tables[0].Rows[7][2]);
            Assert.AreEqual(99, result.Tables[0].Rows[8][2]);
            Assert.AreEqual(1385729.234D, result.Tables[0].Rows[9][2]);
        }

        [TestMethod]
        public void GitIssue240ExceptionBeforeRead()
        {
            // Check the exception and message when trying to get data before calling Read().
            // Using the same as SqlDataReader, making it easier to search for a general solution.
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test10x10.xls")))
            {
                var exception = Assert.Throws<InvalidOperationException>(() =>
                {
                    for (int columnIndex = 0; columnIndex < excelReader.FieldCount; columnIndex++)
                    {
                        string _ = excelReader.GetString(columnIndex);
                    }
                });

                Assert.AreEqual("No data exists for the row/column.", exception.Message);
            }
        }

        [TestMethod]
        public void GitIssue241Simple95()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_224_simple_95.xls")))
            {
                Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft åäö &T&CCenter åäö &D&RRight  åäö &P"), "Header");
                Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft åäö &P&CFooter åäö &P&RRight åäö &D"), "Footer");
            }
        }

        [TestMethod]
        public void GitIssue245CodeNameHoja8()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_45.xls")))
            {
                Assert.AreEqual("Hoja8", reader.CodeName);
            }
        }
        
        [TestMethod]
        public void GitIssue242Password()
        {
            // BIFF8 standard encryption cryptoapi rc4+sha 
            using (var reader = ExcelReaderFactory.CreateBinaryReader(
                Configuration.GetTestWorkbook("Test_git_issue_242_std_rc4_pwd_password.xls"),
                new ExcelReaderConfiguration { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }

            // Pre-BIFF8 xor obfuscation
            using (var reader = ExcelReaderFactory.CreateBinaryReader(
                Configuration.GetTestWorkbook("Test_git_issue_242_xor_pwd_password.xls"),
                new ExcelReaderConfiguration { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual("Password: password", reader.GetString(0));
            }
        }

        [TestMethod]
        public void BinaryThrowsInvalidPassword()
        {
            Assert.Throws<InvalidPasswordException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateBinaryReader(
                    Configuration.GetTestWorkbook("Test_git_issue_242_xor_pwd_password.xls"),
                    new ExcelReaderConfiguration { Password = "wrongpassword" }))
                {
                    reader.Read();
                }
            });

            Assert.Throws<InvalidPasswordException>(() =>
            {
                using (var reader = ExcelReaderFactory.CreateBinaryReader(
                    Configuration.GetTestWorkbook("Test_git_issue_242_xor_pwd_password.xls")))
                {
                    reader.Read();
                }
            });
        }

        [TestMethod]
        public void GitIssue263()
        {
            using (var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test_git_issue_263.xls")))
            {
                var ds = reader.AsDataSet();
                Assert.AreEqual("Economic Inactivity by age\n(Official statistics: not designated as National Statistics)", ds.Tables[1].Rows[3][0]);
            }
        }

        [TestMethod]
        public void GitIssue265BinaryDisposed()
        {
            var stream = Configuration.GetTestWorkbook("Test10x10.xls");
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream))
            {
                var _ = excelReader.AsDataSet();
            }

            Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
        }

        [TestMethod]
        public void BinaryRawLeaveOpen()
        {
            // Verify raw stream is not disposed by the reader
            {
                var stream = Configuration.GetTestWorkbook("biff3.xls");
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration()
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

        [TestMethod]
        public void GitIssue286SstStringHeader()
        {
            // Parse xls with SST containing string split exactly between its header and string data across the BIFF Continue records
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_286_SST.xls")))
            {
                Assert.IsNotNull(reader);
            }
        }
        
        [TestMethod]
        public void GitIssue321MissingEof()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue321.xls")))
            {
                for (int i = 0; i < 7; i++)
                {
                    reader.Read();
                    Assert.IsTrue(string.IsNullOrEmpty(reader.GetString(1)), "Row = " + i);
                }

                reader.Read();
                Assert.AreEqual(" MONETARY AGGREGATES FOR INSTITUTIONAL SECTORS", reader.GetString(1));
            }
        }

        [TestMethod]
        public void GitIssue368Header()
        {
            // This reads a specially crafted XLS which loads in Excel:
            // - Raw BIFF5/8 BIFF stream
            // - Non-standard header with size=6, and version=0
            // - Mixes record identifiers from different BIFF versions:
            // - Uses NUMBER (BIFF3-8) and NUMBER_OLD (BIFF2) records
            // - Uses LABEL (BIFF3-5) and LABEL_OLD (BIFF2) records
            // - Uses RK (BIFF3-5) and INTEGER (BIFF2) records
            // - Uses FORMAT_V23 (BIFF2-3) and FORMAT (BIFF4-8) records

            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_368_header.xls")))
            {
                reader.Read();
                Assert.AreEqual("BIFF2", reader[0]);
                Assert.AreEqual(1234.5678, reader[1]);
                Assert.AreEqual(1234, reader[2]);
                Assert.AreEqual("00.0", reader.GetNumberFormatString(1));
                Assert.AreEqual("00.0", reader.GetNumberFormatString(2));

                reader.Read();
                Assert.AreEqual("BIFF3-5", reader[0]);
                Assert.AreEqual(8765.4321, reader[1]);
                Assert.AreEqual(4321, reader[2]);
                Assert.AreEqual("0000.00", reader.GetNumberFormatString(1));
                Assert.AreEqual("0000.00", reader.GetNumberFormatString(2));
            }
        }

        [TestMethod]
        public void GitIssue368Formats()
        {
            // This reads a BIFF2 XLS worksheet created with Excel 2.0 containing 63 number formats, the maximum allowed by the UI.
            // Excel 2.0/2.1 does not write XF/IXFE records, but writes the FORMAT index as 6 bits in the cell attributes.
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_368_formats.xls")))
            {
                for (var i = 0; i < 42; i++)
                {
                    reader.Read();
                    Assert.AreEqual(i % 10, reader[0]);
                    Assert.AreEqual("\"" + i + "\" 0.00", reader.GetNumberFormatString(0));
                }
            }
        }

        [TestMethod]
        public void GitIssue368Ixfe()
        {
            // This reads a specially crafted XLS which loads in Excel:
            // - BIFF2 worksheet, only BIFF2 records
            // - Uses IXFE records to set format
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_368_ixfe.xls")))
            {
                reader.Read();
                Assert.AreEqual("BIFF2", reader[0]);
                Assert.AreEqual(1234.5678, reader[1]);
                Assert.AreEqual(1234, reader[2]);
                Assert.AreEqual("00.0", reader.GetNumberFormatString(1));
                Assert.AreEqual("00.0", reader.GetNumberFormatString(2));

                reader.Read();
                Assert.AreEqual("BIFF2!", reader[0]);
                Assert.AreEqual(8765.4321, reader[1]);
                Assert.AreEqual(4321, reader[2]);
                Assert.AreEqual("0000.00", reader.GetNumberFormatString(1));
                Assert.AreEqual("0000.00", reader.GetNumberFormatString(2));
            }
        }

        [TestMethod]
        public void GitIssue368LabelXf()
        {
            // This reads a specially crafted XLS which loads in Excel:
            // - BIFF2 worksheet, with mixed version FORMAT records, BIFF3-5 label records and 16 bit XF index
            // - Contains 80 XF records
            // - Excel uses only 6 bits of the BIFF3-5 XF index when present in a BIFF2 worksheet, must use IXFE for >62
            // - Excel 2.0 does not write XF>63, but newer Excels read these records
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_368_label_xf.xls")))
            {
                reader.Read();
                Assert.AreEqual("BIFF3-5 record in BIFF2 worksheet with XF 60", reader[0]);
                Assert.AreEqual("\\A@\\B", reader.GetNumberFormatString(0));

                reader.Read();
                Assert.AreEqual("Same with XF 70 (ignored by Excel)", reader[0]);
                // TODO:
                Assert.AreEqual("General", reader.GetNumberFormatString(0));

                reader.Read();
                Assert.AreEqual("Same with XF 70 via IXFE", reader[0]);
                Assert.AreEqual("\\A@\\B", reader.GetNumberFormatString(0));
            }
        }

        [TestMethod]
        public void GitIssue375IxfeRowMap()
        {
            // This reads a specially crafted XLS which loads in Excel:
            // - 100 rows with IXFE records
            // Verify the internal map of cell offsets used for buffering includes the preceding IXFE records
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_375_ixfe_rowmap.xls")))
            {
                for (var i = 0; i < 100; i++)
                {
                    reader.Read();
                    Assert.AreEqual(1234.0 + i + (i / 10.0), reader[0]);
                    Assert.AreEqual("0.000", reader.GetNumberFormatString(0));
                }
            }
        }

        [TestMethod]
        public void GitIssue382Oom()
        {
            Assert.Throws(typeof(CompoundDocumentException), () =>
            {
                using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_382_oom.xls")))
                {
                    reader.AsDataSet();
                }
            });
        }

        [TestMethod]
        public void GitIssue392Oob()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_392_oob.xls")))
            {
                var result = reader.AsDataSet().Tables[0];

                Assert.AreEqual(10, result.Rows.Count);
                Assert.AreEqual(10, result.Columns.Count);
                Assert.AreEqual("10x10", result.Rows[1][0]);
                Assert.AreEqual("10x27", result.Rows[9][9]);
            }
        }

        [TestMethod(Description = "XF_USED_ATTRIB is not set correctly")]
        public void GitIssue_341_HorizontalAlignment2()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_51.xls")))
            {
                Assert.IsTrue(reader.Read());
                Assert.IsTrue(reader.Read());
                Assert.IsTrue(reader.Read());
                Assert.AreEqual(HorizontalAlignment.Right, reader.GetCellStyle(1).HorizontalAlignment);
            }
        }

        [TestMethod(Description = "Indent is from a style")]
        public void GitIssue_341_FromStyle()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_341_style.xls")))
            {
                Assert.IsTrue(reader.Read());
                Assert.AreEqual(2, reader.GetCellStyle(0).IndentLevel);
            }
        }

        [TestMethod]
        public void MultiCellCustomFormatNotDate()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("customformat_notdate.xls")))
            {
                Assert.IsTrue(reader.Read());
                Assert.AreEqual(60.8, reader.GetValue(1));
                Assert.AreEqual("#,##0.0;\\–#,##0.0;\"–\"", reader.GetNumberFormatString(1));
            }
        }

        [TestMethod]
        public void Test_git_issue_411()
        {
            // This file has two problems: 
            // - has both Book and Workbook compound streams
            // - has no codepage record, encoding specified in font records
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_411.xls")))
            {
                Assert.AreEqual(1, reader.ResultsCount);
                Assert.IsTrue(reader.Read());
                Assert.IsTrue(reader.Read());
                Assert.AreEqual("Универсальный передаточный\nдокумент", reader.GetValue(1));
            }
        }

        [TestMethod]
        public void GitIssue438() 
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_438.xls")))
            {
                reader.Read();

                Assert.AreEqual(new DateTime(1992, 05, 15), reader.GetDateTime(0));
            }
        }

        [Test]
        public void GitIssue_341_Indent()
        {
            int[][] expected =
            {
                new[] { 2, 0, 0 },
                new[] { 2, 0, 0 },
                new[] { 3, 3, 4 },
                new[] { 1, 1, 0 }, // Merged cell
                new[] { 2, 0, 0 },
            };

            int index = 0;
            using (var reader = OpenReader("Test_git_issue_341"))
            {
                while (reader.Read())
                {
                    int[] expectedRow = expected[index];
                    int[] actualRow = new int[reader.FieldCount];
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        actualRow[i] = reader.GetCellStyle(i).IndentLevel;
                    }

                    Assert.AreEqual(expectedRow, actualRow, "Indent level on row '{0}'.", index);

                    index++;
                }
            }
        }

        [Test]
        public void GitIssue_341_HorizontalAlignment()
        {
            HorizontalAlignment[][] expected =
            {
                new[] { HorizontalAlignment.Left, HorizontalAlignment.General, HorizontalAlignment.General },
                new[] { HorizontalAlignment.Distributed, HorizontalAlignment.General, HorizontalAlignment.General },
                new[] { HorizontalAlignment.Left, HorizontalAlignment.Left, HorizontalAlignment.Left },
                new[] { HorizontalAlignment.Left, HorizontalAlignment.Left, HorizontalAlignment.General }, // Merged cell
                new[] { HorizontalAlignment.Left, HorizontalAlignment.General, HorizontalAlignment.General },
            };

            int index = 0;
            using (var reader = OpenReader("Test_git_issue_341"))
            {
                while (reader.Read())
                {
                    HorizontalAlignment[] expectedRow = expected[index];
                    HorizontalAlignment[] actualRow = new HorizontalAlignment[reader.FieldCount];
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        actualRow[i] = reader.GetCellStyle(i).HorizontalAlignment;
                    }

                    Assert.AreEqual(expectedRow, actualRow, "Horizontal alignment on row '{0}'.", index);

                    index++;
                }
            }
        }

        [TestMethod]
        public void GitIssue477_Test_crypto_keylength40()
        {
            // BIFF8 standard encryption cryptoapi rc4+sha with 40bit key
            // Test file from SheetJS project: password_2002_40_basecrypto.xls
            using (var reader = ExcelReaderFactory.CreateBinaryReader(
                Configuration.GetTestWorkbook("Test_git_issue_477_crypto_keylength40.xls"),
                new ExcelReaderConfiguration { Password = "password" }))
            {
                reader.Read();
                Assert.AreEqual(1, reader.GetDouble(0));
 
                reader.Read();
                Assert.AreEqual(2, reader.GetDouble(0));
                Assert.AreEqual(10, reader.GetDouble(1));
            }
        }

        [TestMethod]
        public void GitIssue467_Test_empty_continue_SST()
        {
            // File was modified in a hex editor to include an empty CONTINUE record with only a multi byte flag
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_467_sst_empty_continue.xls")))
            {
                reader.Read();
            }
        }

        [TestMethod]
        public void GitIssue467_Test_emptier_continue_leftover_bytes_SST()
        {
            // File was modified in a hex editor to include an empty CONTINUE record without a multi byte flag
            // followed by a CONTINUE record with multibyte flag and a leftover byte
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_467_empty_continue_leftoverbytes.xls")))
            {
                reader.Read();
            }
        }

        [TestMethod]
        public void GitIssue467_Test_SST_wrong_count()
        {
            // Modified 10x10.xls in a hex editor to specify too many strings in the SST
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_477_sst_wrong_count.xls")))
            {
                reader.Read();
                Assert.AreEqual(10, reader.RowCount);
                Assert.AreEqual(10, reader.FieldCount);
                Assert.AreEqual("col1", reader.GetString(0));
                Assert.AreEqual("col3", reader.GetString(2));
                Assert.AreEqual("col7", reader.GetString(6));

                reader.Read();
                Assert.AreEqual("10x10", reader.GetString(0));

                reader.Read();
                reader.Read();
                reader.Read();
                reader.Read();
                reader.Read();
                reader.Read();
                reader.Read();
                reader.Read();
                reader.Read();
                Assert.AreEqual("10x27", reader.GetString(9));
            }
        }

        [TestMethod]
        public void GitIssue467_Test_SST_zero_count()
        {
            // Modified 10x10.xls in a hex editor to specify zero strings in the SST: Excel doesn't read these
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_477_sst_zero_count.xls")))
            {
                reader.Read();
                Assert.AreEqual(10, reader.RowCount);
                Assert.AreEqual(10, reader.FieldCount);
                Assert.AreEqual(null, reader.GetString(0));
                Assert.AreEqual(null, reader.GetString(2));
                Assert.AreEqual(null, reader.GetString(6));

                reader.Read();
                Assert.AreEqual(null, reader.GetString(0));

                reader.Read();
                reader.Read();
                reader.Read();
                reader.Read();
                reader.Read();
                reader.Read();
                reader.Read();
                reader.Read();
                reader.Read();
                Assert.AreEqual(null, reader.GetString(9));
            }
        }

        [TestMethod]
        public void GitIssue466_BIFF3_Errors()
        {
            using (var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_466_biff3.xls")))
            {
                // First row contains formula errors
                reader.Read();
                Assert.AreEqual(null, reader.GetString(0));
                Assert.AreEqual(CellError.DIV0, reader.GetCellError(0));

                Assert.AreEqual(null, reader.GetString(1));
                Assert.AreEqual(CellError.NA, reader.GetCellError(1));

                Assert.AreEqual(null, reader.GetString(2));
                Assert.AreEqual(CellError.VALUE, reader.GetCellError(2));

                Assert.AreEqual(null, reader.GetString(3));
                Assert.AreEqual(CellError.NAME, reader.GetCellError(3));

                Assert.AreEqual(null, reader.GetString(4));
                Assert.AreEqual(CellError.REF, reader.GetCellError(4));
                
                // Second row contains error constants
                reader.Read();
                Assert.AreEqual(null, reader.GetString(0));
                Assert.AreEqual(CellError.DIV0, reader.GetCellError(0));

                Assert.AreEqual(null, reader.GetString(1));
                Assert.AreEqual(CellError.NA, reader.GetCellError(1));

                Assert.AreEqual(null, reader.GetString(2));
                Assert.AreEqual(CellError.VALUE, reader.GetCellError(2));

                Assert.AreEqual(null, reader.GetString(3));
                Assert.AreEqual(CellError.NAME, reader.GetCellError(3));

                Assert.AreEqual(null, reader.GetString(4));
                Assert.AreEqual(CellError.REF, reader.GetCellError(4));
            }
        }

        [Test]
        public void GitIssue532MulCells()
        {
            using var reader = OpenReader("Test_git_issue_532_mulcells");
            reader.NextResult();
            reader.Read();

            Assert.AreEqual(77, reader.FieldCount);
        }
    }
}
