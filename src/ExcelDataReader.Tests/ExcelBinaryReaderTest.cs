using System.Data;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Tests;

public class ExcelBinaryReaderTest : ExcelTestBase
{
    /// <inheritdoc />
    protected override DateTime GitIssue82TodayDate => new(2009, 6, 28);

    [Test]
    public void GitIssue70ExcelBinaryReaderTryConvertOADateTimeFormula()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_70_ExcelBinaryReader_tryConvertOADateTime _convert_dates.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds, Is.Not.Null);

        var date = ds.Tables[0].Rows[1].ItemArray[0];

        Assert.That(date, Is.EqualTo(new DateTime(2014, 01, 01)));
    }

    [Test]
    public void GitIssue51ReadCellLabel()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_51.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds, Is.Not.Null);

        var value = ds.Tables[0].Rows[0].ItemArray[1];

        Assert.That(value, Is.EqualTo("Monetary aggregates (R millions)"));
    }

    [Test]
    public void GitIssue45()
    {
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_45.xls"));
        do
        {
            while (reader.Read())
            {
            }
        }
        while (reader.NextResult());
    }

    [Test]
    public void Issue1155311570FatIssueOffset()
    {
        void DoTestFatStreamIssue(string sheetId)
        {
            string filePath;
            using (var excelReader1 = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook(sheetId))) // Works.
            {
                filePath = Configuration.GetTestWorkbookPath(sheetId);
                Assert.That(excelReader1, Is.Not.Null);
            }

            using (var ms1 = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelReader2 = ExcelReaderFactory.CreateBinaryReader(ms1)) // Works!
                Assert.That(excelReader2, Is.Not.Null);

            var bytes = File.ReadAllBytes(filePath);
            using var ms2 = new MemoryStream(bytes);
            using var excelReader3 = ExcelReaderFactory.CreateBinaryReader(ms2); // Did not work, but does now
            Assert.That(excelReader3, Is.Not.Null);
        }

        void DoTestFatStreamIssueType2(string sheetId)
        {
            var filePath = Configuration.GetTestWorkbookPath(sheetId);

            using Stream stream = new MemoryStream(File.ReadAllBytes(filePath));
            using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            Assert.DoesNotThrow(() => excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration));
        }

        DoTestFatStreamIssue("Test_Issue_11553_FAT.xls");
        DoTestFatStreamIssueType2("Test_Issue_11570_FAT_1.xls");
        DoTestFatStreamIssueType2("Test_Issue_11570_FAT_2.xls");
    }

    /*[Test]
    public void Test_SSRS()
    {
        IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_SSRS"));
        DataSet result = excelReader.AsDataSet();
        excelReader.Close();
    }*/

    [Test]
    public void FailTest()
    {
        var exception = Assert.Throws<HeaderException>(() =>
        {
            using (ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestFail_Binary.xls")))
            {
            }
        });

        Assert.That(exception.Message, Is.EqualTo("Invalid file signature."));
    }

    [Test]
    public void TestOpenOffice()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_OpenOffice.xls"));
        AssertUtilities.DoOpenOfficeTest(excelReader);
    }

    /// <summary>
    /// Issue 11 - OpenOffice files were skipping the first row if IsFirstRowAsColumnNames = false.
    /// </summary>
    [Test]
    public void GitIssue11OpenOfficeRowCount()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_OpenOffice.xls"));
        var dataSet = excelReader.AsDataSet(Configuration.NoColumnNamesConfiguration);
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(34));
    }
    
    [Test]
    public void UncalculatedTest()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Uncalculated.xls"));
        var dataSet = excelReader.AsDataSet();
        Assert.That(dataSet, Is.Not.Null);
        Assert.That(dataSet.Tables.Count, Is.Not.EqualTo(0));
        var table = dataSet.Tables[0];
        Assert.That(table, Is.Not.Null);

        Assert.That(table.Rows[1][0].ToString(), Is.EqualTo("1"));
        Assert.That(table.Rows[1][2].ToString(), Is.EqualTo("3"));
        Assert.That(table.Rows[1][4].ToString(), Is.EqualTo("3"));
    }

    [Test]
    public void Issue11570Excel2013()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11570_Excel2013.xls"));
        var dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(2));
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(5));

        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo("1.1.1.2"));
        Assert.That(dataSet.Tables[0].Rows[0][1], Is.EqualTo(10d));

        Assert.That(dataSet.Tables[0].Rows[1][0], Is.EqualTo("1.1.1.15"));
        Assert.That(dataSet.Tables[0].Rows[1][1], Is.EqualTo(3d));

        Assert.That(dataSet.Tables[0].Rows[2][0], Is.EqualTo("2.1.2.23"));
        Assert.That(dataSet.Tables[0].Rows[2][1], Is.EqualTo(14d));

        Assert.That(dataSet.Tables[0].Rows[3][0], Is.EqualTo("2.1.2.31"));
        Assert.That(dataSet.Tables[0].Rows[3][1], Is.EqualTo(2d));

        Assert.That(dataSet.Tables[0].Rows[4][0], Is.EqualTo("2.8.7.30"));
        Assert.That(dataSet.Tables[0].Rows[4][1], Is.EqualTo(2d));
    }

    [Test]
    public void Issue11572CodePage()
    {
        // This test was skipped for a long time as it produced: "System.NotSupportedException : No data is available for encoding 27651."
        // Upon revisiting the underlying cause appears to be fixed
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11572_CodePage.xls"));
        Assert.DoesNotThrow(() => excelReader.AsDataSet());
    }

    /// <summary>
    /// Not fixed yet.
    /// </summary>
    [Test]
    public void Issue11545NoIndex()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11545_NoIndex.xls"));
        var dataSet = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);

        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo("CI2229         "));
        Assert.That(dataSet.Tables[0].Rows[0][6], Is.EqualTo("12069E01018A1  "));
        Assert.That(dataSet.Tables[0].Rows[0][8], Is.EqualTo(new DateTime(2012, 03, 01)));
    }

    [Test]
    public void Issue11642ValuesNotLoaded()
    {
        // Excel.Log.Log.InitializeWith<Log4NetLog>();
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11642_ValuesNotLoaded.xls"));
        var dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[2].Rows[29][1].ToString(), Is.EqualTo("431113*"));
        Assert.That(dataSet.Tables[2].Rows[36][1].ToString(), Is.EqualTo("024807"));
        Assert.That(dataSet.Tables[2].Rows[53][1].ToString(), Is.EqualTo("160019"));
    }

    [Test]
    public void Issue11636BiffStream()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11636_BiffStream.xls"));
        var dataSet = excelReader.AsDataSet();

        // check a couple of values
        Assert.That(dataSet.Tables[0].Rows[9][0], Is.EqualTo("SP011"));
        Assert.That(dataSet.Tables[0].Rows[32][11], Is.EqualTo(9.9));
        Assert.That(dataSet.Tables[1].Rows[27][12], Is.EqualTo(78624.44));
    }

    /// <summary>
    /// Not fixed yet
    /// The problem occurs with unseekable stream and logic related to minifat that uses seek
    /// It should probably only use seek if it needs to go backwards, I think at the moment it uses seek all the time
    /// which is probably not good for performance.
    /// </summary>
    [Test]
    [Ignore("Not fixed yet")]
    public void Issue1163911644ForwardOnlyStream()
    {
        // Excel.Log.Log.InitializeWith<Log4NetLog>();
        using var stream = Configuration.GetTestWorkbook("Test_OpenOffice");
        using var forwardStream = SeekErrorMemoryStream.CreateFromStream(stream);
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(forwardStream);
        Assert.DoesNotThrow(() => excelReader.AsDataSet());
    }

    /// <summary>
    /// Not fixed yet
    /// The problem occurs with unseekable stream and logic related to minifat that uses seek
    /// It should probably only use seek if it needs to go backwards, I think at the moment it uses seek all the time
    /// which is probably not good for performance.
    /// </summary>
    [Test]
    public void Issue12556Corrupt()
    {
        Assert.Throws<CompoundDocumentException>(() =>
        {
            // Excel.Log.Log.InitializeWith<Log4NetLog>();
            using var forwardStream = Configuration.GetTestWorkbook("Test_Issue_12556_corrupt.xls");
            using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(forwardStream);
            Assert.DoesNotThrow(() => excelReader.AsDataSet());
        });
    }

    /// <summary>
    /// Some spreadsheets were crashing with index out of range error (from SSRS).
    /// </summary>
    [Test]
    public void TestIssue11818OutOfRange()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Issue_11818_OutOfRange.xls"));
        var dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[10][0], Is.EqualTo("Total Revenue"));
    }

    [Test]
    public void TestIssue111NoRowRecords()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_111_NoRowRecords.xls"));
        var dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables.Count, Is.EqualTo(1));
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(12));
        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(14));

        Assert.That(dataSet.Tables[0].Rows[7][0], Is.EqualTo(2015.0));
    }

    [Test]
    public void TestGitIssue145()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_145.xls"));
        excelReader.Read();
        excelReader.Read();
        excelReader.Read();

        string value = excelReader.GetString(3);

        Assert.That(value, Is.EqualTo("Japanese Government Bonds held by the Bank of Japan"));
    }

    [Test]
    public void TestGitIssue152SheetNameUtf16LeCompressed()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_152.xls"));
        var dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[0].TableName, Is.EqualTo("åäöñ"));
    }

    [Test]
    public void TestGitIssue152CellUtf16LeCompressed()
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_152.xls"));
        var dataSet = excelReader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo("åäöñ"));
    }

    [Test]
    public void GitIssue158()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_158.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds, Is.Not.Null);

        var date = ds.Tables[0].Rows[3].ItemArray[2];

        Assert.That(date, Is.EqualTo(new DateTime(2016, 09, 10)));
    }

    [Test]
    public void GitIssue173()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_173.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds, Is.Not.Null);
        Assert.That(ds.Tables.Count, Is.EqualTo(40));
    }

    [Test]
    public void ReadWriteProtectedStructureUsingStandardEncryption()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("protectedsheet-xxx.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds, Is.Not.Null);
        Assert.That(ds.Tables[0].Rows[0][0], Is.EqualTo("x"));
        Assert.That(ds.Tables[0].Rows[1][0], Is.EqualTo(1.4));
    }

    [Test]
    public void TestIncludeTableWithOnlyImage()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("TestTableOnlyImage_x01oct2016.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds, Is.Not.Null);
        Assert.That(ds.Tables.Count, Is.EqualTo(4));
    }

    [Test]
    public void AllowFfffAsByteOrder()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_InvalidByteOrderValueInHeader.xls"));
        int tableCount = 0;
        do
        {
            while (excelReader.Read())
            {
            }

            tableCount++;
        }
        while (excelReader.NextResult());

        Assert.That(tableCount, Is.EqualTo(454));
    }

    [Test]
    public void HandleRowBlocksWithOutOfOrderCells()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("AllColumnsNotReadInHiddenTable.xls"));
        var ds = excelReader.AsDataSet();

        object[] expected = { "21/09/2015", 1187.5282349881188, 650.8582749049624, 1361.7209439645526, 321.74647548613916, 369.48879457369037 };

        Assert.That(ds.Tables[1].Rows.Count, Is.EqualTo(51));
        Assert.That(ds.Tables[1].Rows[1].ItemArray, Is.EqualTo(expected));
    }

    [Test]
    public void HandleRowBlocksWithDifferentNumberOfColumnsAndInvalidDimensions()
    {
        // http://www.ine.cl/canales/chile_estadistico/estadisticas_economicas/edificacion/archivos/xls/edificacion_totalpais_seriehistorica_enero_2017.xls
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("RowWithDifferentNumberOfColumns.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds.Tables[0].Columns.Count, Is.EqualTo(256));
    }

    [Test]
    public void IfNoDimensionDetermineFieldCountByProcessingAllCellColumnIndexes()
    {
        // This xls file has a row record with 256 columns but only values for 6.
        // This test was created when ExcelDataReader incorrectly dropped 8
        // bits off the dimensions' LastColumn in BIFF8 files and relied
        // on scanning to come up with 6 columns. The test was changed to
        // assume valid dimensions:
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_145.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds.Tables[0].Columns.Count, Is.EqualTo(5));
    }

    [Test]
    public void Row1217NotRead()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Row1217NotRead.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(
            ds.Tables[0].Rows[1216].ItemArray, 
            Is.EqualTo(new object[] 
            {
                DBNull.Value,
                "Año",
                "Mes",
                DBNull.Value, 
                "Índice",
                "Variación Mensual",
                "Variación Acumulada",
                "Variación en 12 Meses",
                "Incidencia Mensual",
                "Incidencia Acumulada", 
                "Incidencia a 12 Meses",
            }));
    }

    [Test]
    public void StringContinuationAfterCharacterData()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("StringContinuationAfterCharacterData.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds.Tables[0].Rows[3][2], Is.EqualTo("商業動態統計速報-平成29年2月分-  統計表"));
        Assert.That(ds.Tables[0].Rows[4][2], Is.EqualTo("Preliminary Report on the Current Survey of Commerce  ( February,2017 )　Statistics Tables"));
        Assert.That(ds.Tables[1].Rows[18][9], Is.EqualTo("\nWholesale"));
    }

    [TestCase]
    public void Biff3IsSupported()
    {
        using var stream = Configuration.GetTestWorkbook("biff3.xls");
        using var reader = ExcelReaderFactory.CreateBinaryReader(stream);
        reader.AsDataSet();
    }

    [TestCase]
    public void GitIssue5()
    {
        using var stream = Configuration.GetTestWorkbook("Test_git_issue_5.xls");
        Assert.Throws<CompoundDocumentException>(() => ExcelReaderFactory.CreateBinaryReader(stream));
    }

    [TestCase]
    public void Issue2InvalidDimensionRecord()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_2.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds.Tables[0].Rows[0].ItemArray, Is.EqualTo(new[] { "A1", "B1" }));
    }

    [TestCase]
    public void ExcelLibraryNonContinuousMiniStream()
    {
        // Verify the output from the sample code for the ExcelLibrary package parses
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("ExcelLibrary_newdoc.xls"));
        Assert.DoesNotThrow(() => excelReader.AsDataSet());
    }

    [TestCase]
    public void GitIssue184AdditionalFatSectors()
    {
        // Big spreadsheets have additional sectors beyond the header with FAT contents
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("GitIssue_184_FATSectors.xls"));
        DataSet ds = null;
        Assert.DoesNotThrow(() => ds = excelReader.AsDataSet());
        Assert.That(ds.Tables.Count, Is.EqualTo(12));
        Assert.That(ds.Tables[0].TableName, Is.EqualTo("DATAS (12)"));
        Assert.That(ds.Tables[11].TableName, Is.EqualTo("DATAS (5)"));
    }

    [Test]
    public void RowContentSpreadOverMultipleBlocks()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_217.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds.Tables[2].Rows[10].ItemArray, Is.EqualTo(new object[] { "REX GESAMT      ", 484.7929, 142.1032, -0.1656, 5.0315225293000001, 5.0398685515999997, 37.5344725251 }).AsCollection);
    }

    [Test]
    public void GitIssue231NoCodePage()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_231_NoCodePage.xls"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds.Tables[0].Columns.Count, Is.EqualTo(11));
        Assert.That(ds.Tables[0].Rows.Count, Is.EqualTo(5));
    }

    [Test]
    public void As3XlsBiff2()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF2.xls"));
        DataSet result = excelReader.AsDataSet();
        TestAs3Xls(result);
    }

    [Test]
    public void As3XlsBiff3()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF3.xls"));
        DataSet result = excelReader.AsDataSet();
        TestAs3Xls(result);
    }

    [Test]
    public void As3XlsBiff4()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF4.xls"));
        DataSet result = excelReader.AsDataSet();
        TestAs3Xls(result);
    }

    [Test]
    public void As3XlsBiff5()
    {
        using var excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("as3xls_BIFF5.xls"));
        DataSet result = excelReader.AsDataSet();
        TestAs3Xls(result);
    }

    [Test]
    public void GitIssue240ExceptionBeforeRead()
    {
        // Check the exception and message when trying to get data before calling Read().
        // Using the same as SqlDataReader, making it easier to search for a general solution.
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test10x10.xls"));
        var exception = Assert.Throws<InvalidOperationException>(() =>
        {
            for (int columnIndex = 0; columnIndex < excelReader.FieldCount; columnIndex++)
            {
                _ = excelReader.GetString(columnIndex);
            }
        });

        Assert.That(exception.Message, Is.EqualTo("No data exists for the row/column."));
    }

    [Test]
    public void GitIssue241Simple95()
    {
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_224_simple_95.xls"));
        Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft åäö &T&CCenter åäö &D&RRight  åäö &P"), "Header");
        Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft åäö &P&CFooter åäö &P&RRight åäö &D"), "Footer");
    }

    [Test]
    public void GitIssue245CodeNameHoja8()
    {
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_45.xls"));
        Assert.That(reader.CodeName, Is.EqualTo("Hoja8"));
    }
    
    [Test]
    public void GitIssue242Password()
    {
        // BIFF8 standard encryption cryptoapi rc4+sha 
        using (var reader = ExcelReaderFactory.CreateBinaryReader(
            Configuration.GetTestWorkbook("Test_git_issue_242_std_rc4_pwd_password.xls"),
            new ExcelReaderConfiguration { Password = "password" }))
        {
            reader.Read();
            Assert.That(reader.GetString(0), Is.EqualTo("Password: password"));
        }

        // Pre-BIFF8 xor obfuscation
        using (var reader = ExcelReaderFactory.CreateBinaryReader(
            Configuration.GetTestWorkbook("Test_git_issue_242_xor_pwd_password.xls"),
            new ExcelReaderConfiguration { Password = "password" }))
        {
            reader.Read();
            Assert.That(reader.GetString(0), Is.EqualTo("Password: password"));
        }
    }

    [Test]
    public void BinaryThrowsInvalidPassword()
    {
        Assert.Throws<InvalidPasswordException>(() =>
        {
            using var reader = ExcelReaderFactory.CreateBinaryReader(
                Configuration.GetTestWorkbook("Test_git_issue_242_xor_pwd_password.xls"),
                new ExcelReaderConfiguration { Password = "wrongpassword" });
            reader.Read();
        });

        Assert.Throws<InvalidPasswordException>(() =>
        {
            using var reader = ExcelReaderFactory.CreateBinaryReader(
                Configuration.GetTestWorkbook("Test_git_issue_242_xor_pwd_password.xls"));
            reader.Read();
        });
    }

    [Test]
    public void GitIssue263()
    {
        using var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test_git_issue_263.xls"));
        var ds = reader.AsDataSet();
        Assert.That(ds.Tables[1].Rows[3][0], Is.EqualTo("Economic Inactivity by age\n(Official statistics: not designated as National Statistics)"));
    }

    [Test]
    public void GitIssue265BinaryDisposed()
    {
        var stream = Configuration.GetTestWorkbook("Test10x10.xls");
        using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream))
        {
            _ = excelReader.AsDataSet();
        }

        Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
    }

    [Test]
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
                _ = excelReader.AsDataSet();
            }

            stream.Seek(0, SeekOrigin.Begin);
            stream.ReadByte();
            stream.Dispose();
        }
    }

    [Test]
    public void GitIssue286SstStringHeader()
    {
        // Parse xls with SST containing string split exactly between its header and string data across the BIFF Continue records
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_286_SST.xls"));
        Assert.That(reader, Is.Not.Null);
    }
    
    [Test]
    public void GitIssue321MissingEof()
    {
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue321.xls"));
        for (int i = 0; i < 7; i++)
        {
            reader.Read();
            Assert.That(string.IsNullOrEmpty(reader.GetString(1)), Is.True, "Row = " + i);
        }

        reader.Read();
        Assert.That(reader.GetString(1), Is.EqualTo(" MONETARY AGGREGATES FOR INSTITUTIONAL SECTORS"));
    }

    [Test]
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
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_368_header.xls"));
        reader.Read();
        Assert.That(reader[0], Is.EqualTo("BIFF2"));
        Assert.That(reader[1], Is.EqualTo(1234.5678));
        Assert.That(reader[2], Is.EqualTo(1234));
        Assert.That(reader.GetNumberFormatString(1), Is.EqualTo("00.0"));
        Assert.That(reader.GetNumberFormatString(2), Is.EqualTo("00.0"));

        reader.Read();
        Assert.That(reader[0], Is.EqualTo("BIFF3-5"));
        Assert.That(reader[1], Is.EqualTo(8765.4321));
        Assert.That(reader[2], Is.EqualTo(4321));
        Assert.That(reader.GetNumberFormatString(1), Is.EqualTo("0000.00"));
        Assert.That(reader.GetNumberFormatString(2), Is.EqualTo("0000.00"));
    }

    [Test]
    public void GitIssue368Formats()
    {
        // This reads a BIFF2 XLS worksheet created with Excel 2.0 containing 63 number formats, the maximum allowed by the UI.
        // Excel 2.0/2.1 does not write XF/IXFE records, but writes the FORMAT index as 6 bits in the cell attributes.
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_368_formats.xls"));
        for (var i = 0; i < 42; i++)
        {
            reader.Read();
            Assert.That(reader[0], Is.EqualTo(i % 10));
            Assert.That(reader.GetNumberFormatString(0), Is.EqualTo("\"" + i + "\" 0.00"));
        }
    }

    [Test]
    public void GitIssue368Ixfe()
    {
        // This reads a specially crafted XLS which loads in Excel:
        // - BIFF2 worksheet, only BIFF2 records
        // - Uses IXFE records to set format
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_368_ixfe.xls"));
        reader.Read();
        Assert.That(reader[0], Is.EqualTo("BIFF2"));
        Assert.That(reader[1], Is.EqualTo(1234.5678));
        Assert.That(reader[2], Is.EqualTo(1234));
        Assert.That(reader.GetNumberFormatString(1), Is.EqualTo("00.0"));
        Assert.That(reader.GetNumberFormatString(2), Is.EqualTo("00.0"));

        reader.Read();
        Assert.That(reader[0], Is.EqualTo("BIFF2!"));
        Assert.That(reader[1], Is.EqualTo(8765.4321));
        Assert.That(reader[2], Is.EqualTo(4321));
        Assert.That(reader.GetNumberFormatString(1), Is.EqualTo("0000.00"));
        Assert.That(reader.GetNumberFormatString(2), Is.EqualTo("0000.00"));
    }

    [Test]
    public void GitIssue368LabelXf()
    {
        // This reads a specially crafted XLS which loads in Excel:
        // - BIFF2 worksheet, with mixed version FORMAT records, BIFF3-5 label records and 16 bit XF index
        // - Contains 80 XF records
        // - Excel uses only 6 bits of the BIFF3-5 XF index when present in a BIFF2 worksheet, must use IXFE for >62
        // - Excel 2.0 does not write XF>63, but newer Excels read these records
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_368_label_xf.xls"));
        reader.Read();
        Assert.That(reader[0], Is.EqualTo("BIFF3-5 record in BIFF2 worksheet with XF 60"));
        Assert.That(reader.GetNumberFormatString(0), Is.EqualTo("\\A@\\B"));

        reader.Read();
        Assert.That(reader[0], Is.EqualTo("Same with XF 70 (ignored by Excel)"));
        
        // TODO:
        Assert.That(reader.GetNumberFormatString(0), Is.EqualTo("General"));

        reader.Read();
        Assert.That(reader[0], Is.EqualTo("Same with XF 70 via IXFE"));
        Assert.That(reader.GetNumberFormatString(0), Is.EqualTo("\\A@\\B"));
    }

    [Test]
    public void GitIssue375IxfeRowMap()
    {
        // This reads a specially crafted XLS which loads in Excel:
        // - 100 rows with IXFE records
        // Verify the internal map of cell offsets used for buffering includes the preceding IXFE records
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_375_ixfe_rowmap.xls"));
        for (var i = 0; i < 100; i++)
        {
            reader.Read();
            Assert.That(reader[0], Is.EqualTo(1234.0 + i + (i / 10.0)));
            Assert.That(reader.GetNumberFormatString(0), Is.EqualTo("0.000"));
        }
    }

    [Test]
    public void GitIssue382Oom()
    {
        Assert.Throws(typeof(CompoundDocumentException), () =>
        {
            using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_382_oom.xls"));
            reader.AsDataSet();
        });
    }

    [Test]
    public void GitIssue392Oob()
    {
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_392_oob.xls"));
        var result = reader.AsDataSet().Tables[0];

        Assert.That(result.Rows.Count, Is.EqualTo(10));
        Assert.That(result.Columns.Count, Is.EqualTo(10));
        Assert.That(result.Rows[1][0], Is.EqualTo("10x10"));
        Assert.That(result.Rows[9][9], Is.EqualTo("10x27"));
    }

    [Test(Description = "XF_USED_ATTRIB is not set correctly")]
    public void GitIssue_341_HorizontalAlignment2()
    {
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_Git_Issue_51.xls"));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetCellStyle(1).HorizontalAlignment, Is.EqualTo(HorizontalAlignment.Right));
    }

    [Test(Description = "Indent is from a style")]
    public void GitIssue_341_FromStyle()
    {
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_341_style.xls"));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetCellStyle(0).IndentLevel, Is.EqualTo(2));
    }

    [Test]
    public void MultiCellCustomFormatNotDate()
    {
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("customformat_notdate.xls"));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetValue(1), Is.EqualTo(60.8));
        Assert.That(reader.GetNumberFormatString(1), Is.EqualTo("#,##0.0;\\–#,##0.0;\"–\""));
    }

    [Test]
    public void Test_git_issue_411()
    {
        // This file has two problems: 
        // - has both Book and Workbook compound streams
        // - has no codepage record, encoding specified in font records
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_411.xls"));
        Assert.That(reader.ResultsCount, Is.EqualTo(1));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetValue(1), Is.EqualTo("Универсальный передаточный\nдокумент"));
    }

    [Test]
    public void GitIssue438() 
    {
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_438.xls"));
        reader.Read();

        Assert.That(reader.GetDateTime(0), Is.EqualTo(new DateTime(1992, 05, 15)));
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
        using var reader = OpenReader("Test_git_issue_341");
        while (reader.Read())
        {
            int[] expectedRow = expected[index];
            int[] actualRow = new int[reader.FieldCount];
            for (int i = 0; i < reader.FieldCount; i++)
            {
                actualRow[i] = reader.GetCellStyle(i).IndentLevel;
            }

            Assert.That(actualRow, Is.EqualTo(expectedRow), $"Indent level on row '{index}'.");

            index++;
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
        using var reader = OpenReader("Test_git_issue_341");
        while (reader.Read())
        {
            HorizontalAlignment[] expectedRow = expected[index];
            HorizontalAlignment[] actualRow = new HorizontalAlignment[reader.FieldCount];
            for (int i = 0; i < reader.FieldCount; i++)
            {
                actualRow[i] = reader.GetCellStyle(i).HorizontalAlignment;
            }

            Assert.That(actualRow, Is.EqualTo(expectedRow), $"Horizontal alignment on row '{index}'.");

            index++;
        }
    }

    [Test]
    public void GitIssue477_Test_crypto_keylength40()
    {
        // BIFF8 standard encryption cryptoapi rc4+sha with 40bit key
        // Test file from SheetJS project: password_2002_40_basecrypto.xls
        using var reader = ExcelReaderFactory.CreateBinaryReader(
            Configuration.GetTestWorkbook("Test_git_issue_477_crypto_keylength40.xls"),
            new ExcelReaderConfiguration { Password = "password" });
        reader.Read();
        Assert.That(reader.GetDouble(0), Is.EqualTo(1));

        reader.Read();
        Assert.That(reader.GetDouble(0), Is.EqualTo(2));
        Assert.That(reader.GetDouble(1), Is.EqualTo(10));
    }

    [Test]
    public void GitIssue467_Test_empty_continue_SST()
    {
        // File was modified in a hex editor to include an empty CONTINUE record with only a multi byte flag
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_467_sst_empty_continue.xls"));
        reader.Read();
    }

    [Test]
    public void GitIssue467_Test_emptier_continue_leftover_bytes_SST()
    {
        // File was modified in a hex editor to include an empty CONTINUE record without a multi byte flag
        // followed by a CONTINUE record with multibyte flag and a leftover byte
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_467_empty_continue_leftoverbytes.xls"));
        reader.Read();
    }

    [Test]
    public void GitIssue467_Test_SST_wrong_count()
    {
        // Modified 10x10.xls in a hex editor to specify too many strings in the SST
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_477_sst_wrong_count.xls"));
        reader.Read();
        Assert.That(reader.RowCount, Is.EqualTo(10));
        Assert.That(reader.FieldCount, Is.EqualTo(10));
        Assert.That(reader.GetString(0), Is.EqualTo("col1"));
        Assert.That(reader.GetString(2), Is.EqualTo("col3"));
        Assert.That(reader.GetString(6), Is.EqualTo("col7"));

        reader.Read();
        Assert.That(reader.GetString(0), Is.EqualTo("10x10"));

        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        Assert.That(reader.GetString(9), Is.EqualTo("10x27"));
    }

    [Test]
    public void GitIssue467_Test_SST_zero_count()
    {
        // Modified 10x10.xls in a hex editor to specify zero strings in the SST: Excel doesn't read these
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_477_sst_zero_count.xls"));
        reader.Read();
        Assert.That(reader.RowCount, Is.EqualTo(10));
        Assert.That(reader.FieldCount, Is.EqualTo(10));
        Assert.That(reader.GetString(0), Is.EqualTo(null));
        Assert.That(reader.GetString(2), Is.EqualTo(null));
        Assert.That(reader.GetString(6), Is.EqualTo(null));

        reader.Read();
        Assert.That(reader.GetString(0), Is.EqualTo(null));

        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        Assert.That(reader.GetString(9), Is.EqualTo(null));
    }

    [Test]
    public void GitIssue466_BIFF3_Errors()
    {
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue_466_biff3.xls"));
        
        // First row contains formula errors
        reader.Read();
        Assert.That(reader.GetString(0), Is.EqualTo(null));
        Assert.That(reader.GetCellError(0), Is.EqualTo(CellError.DIV0));

        Assert.That(reader.GetString(1), Is.EqualTo(null));
        Assert.That(reader.GetCellError(1), Is.EqualTo(CellError.NA));

        Assert.That(reader.GetString(2), Is.EqualTo(null));
        Assert.That(reader.GetCellError(2), Is.EqualTo(CellError.VALUE));

        Assert.That(reader.GetString(3), Is.EqualTo(null));
        Assert.That(reader.GetCellError(3), Is.EqualTo(CellError.NAME));

        Assert.That(reader.GetString(4), Is.EqualTo(null));
        Assert.That(reader.GetCellError(4), Is.EqualTo(CellError.REF));

        // Second row contains error constants
        reader.Read();
        Assert.That(reader.GetString(0), Is.EqualTo(null));
        Assert.That(reader.GetCellError(0), Is.EqualTo(CellError.DIV0));

        Assert.That(reader.GetString(1), Is.EqualTo(null));
        Assert.That(reader.GetCellError(1), Is.EqualTo(CellError.NA));

        Assert.That(reader.GetString(2), Is.EqualTo(null));
        Assert.That(reader.GetCellError(2), Is.EqualTo(CellError.VALUE));

        Assert.That(reader.GetString(3), Is.EqualTo(null));
        Assert.That(reader.GetCellError(3), Is.EqualTo(CellError.NAME));

        Assert.That(reader.GetString(4), Is.EqualTo(null));
        Assert.That(reader.GetCellError(4), Is.EqualTo(CellError.REF));
    }

    [Test]
    public void GitIssue532MulCells()
    {
        using var reader = OpenReader("Test_git_issue_532_mulcells");
        reader.NextResult();
        reader.Read();

        Assert.That(reader.FieldCount, Is.EqualTo(77));
    }

    [Test]
    public void GitIssue624MissingBOFInWorksheet()
    {
        using var reader = OpenReader("Test_git_issue624");
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();
        reader.Read();

        List<string> row = new();
        for (int i = 0; i < reader.FieldCount; i++)
        {
            row.Add(reader.GetString(i));
        }

        Assert.That(row, Is.EqualTo(new[] { "Transref. AM", "NEWM/CANC", "Rel. transref.", "Portfolio ID  AM", "Portfolio ID KVG", "Portfolio name", "BUY/SELL", "OPEP/CLOP", "Quantity", "Instr. ID Type", "Financial instrument ID   ", "Financial instrument name  ", "Unique Product Identifier (UPI)", "Unique Trade Identifier (UTI)", "Price   ", "Strike", "Tic size", "Tic value", "Contract size", "CCY", "Fees", "Clearing amount", "Trade date", "Maturity date", "Sett. CCY", "Place of trade", "Execution Broker ID type", "Execution Broker ID", "Execution Broker Name", "Clearing Broker ID type", "Clearing Broker ID", "Clearing Broker Name", "CCP ID type", "CCP ID", "CCP Name", "Underlying name", "Underlying ISIN", "Put/Call", "Exercise Type", "Settlement Type", "Execution timestamp UTC (Date/Time)", "Confirmation timestamp UTC (Date/Time)", "Confirmation means", "Clearing timestamp UTC (Date/Time)", "Valuta/Settlement Date OTC-Option", "OTC Derivative ISIN", "Order transmission timestamp UTC (Date/Time)" }));
    }

    [Test]
    public void GitIssue609()
    {
        using var reader = OpenReader("Test_git_issue_609");
        reader.Read();
        reader.Read();
        Assert.That(reader.GetString(0), Is.EqualTo("Data di stampa:"));
    }

    [Test]
    public void GitIssue578()
    {
        using var reader = ExcelReaderFactory.CreateBinaryReader(Configuration.GetTestWorkbook("Test_git_issue578.xls"));

        reader.Read();
        var values = new object[reader.FieldCount];
        reader.GetValues(values);
        var values2 = new object[reader.FieldCount + 1];
        reader.GetValues(values2);
        var values3 = new object[reader.FieldCount - 1];
        reader.GetValues(values3);

        Assert.That(values, Is.EqualTo(new object[] { 1, 2, 3, 4, 5 }));
        Assert.That(values2, Is.EqualTo(new object[] { 1, 2, 3, 4, 5, null }));
        Assert.That(values3, Is.EqualTo(new object[] { 1, 2, 3, 4 }));
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

    protected override IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null)
    {
        return ExcelReaderFactory.CreateBinaryReader(stream, configuration);
    }

    protected override Stream OpenStream(string name)
    {
        return Configuration.GetTestWorkbook(name + ".xls");
    }

    private static void TestAs3Xls(DataSet result)
    {
        Assert.That(result.Tables[0].Rows[0][0], Is.EqualTo(1));
        Assert.That(result.Tables[0].Rows[0][1], Is.EqualTo("Hi"));
        Assert.That(result.Tables[0].Rows[0][2], Is.EqualTo(10.22D));
        Assert.That(result.Tables[0].Rows[0][3], Is.EqualTo(14.754317602356753D));
        Assert.That(result.Tables[0].Rows[0][4], Is.EqualTo(21.04107572533686D));

        Assert.That(result.Tables[0].Rows[1][0], Is.EqualTo(2));
        Assert.That(result.Tables[0].Rows[1][1], Is.EqualTo("How"));
        Assert.That(result.Tables[0].Rows[1][2], Is.EqualTo(new DateTime(2007, 2, 22)));

        Assert.That(result.Tables[0].Rows[2][0], Is.EqualTo(3));
        Assert.That(result.Tables[0].Rows[2][1], Is.EqualTo("are"));
        Assert.That(result.Tables[0].Rows[2][2], Is.EqualTo(new DateTime(2002, 1, 19)));

        Assert.That(result.Tables[0].Rows[3][2], Is.EqualTo("Saturday"));
        Assert.That(result.Tables[0].Rows[4][2], Is.EqualTo(0.33000000000000002D));
        Assert.That(result.Tables[0].Rows[5][2], Is.EqualTo(19));
        Assert.That(result.Tables[0].Rows[6][2], Is.EqualTo("Goog"));
        Assert.That(result.Tables[0].Rows[7][2], Is.EqualTo(12.19D));
        Assert.That(result.Tables[0].Rows[8][2], Is.EqualTo(99));
        Assert.That(result.Tables[0].Rows[9][2], Is.EqualTo(1385729.234D));
    }
}
