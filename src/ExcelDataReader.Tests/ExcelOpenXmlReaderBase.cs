using System.Data;

namespace ExcelDataReader.Tests;

public abstract class ExcelOpenXmlReaderBase : ExcelTestBase
{
    protected virtual bool UsesBinaryWorksheet => false;

    protected virtual bool UsesStrictWorksheet => false;

    [Test]
    public void Issue14_InvalidOADate()
    {
        using var excelReader = OpenReader("Issue14_InvalidOADate");
        var dataSet = excelReader.AsDataSet();

        // Test out of range double formatted as date returns double
        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo(1000000000000D));
    }

    [Test]
    public void Issue364()
    {
        using var reader = OpenReader("Issue364");
        Assert.That(reader.RowCount, Is.EqualTo(1));
        reader.Read();

        Assert.That(reader.GetNumberFormatIndex(0), Is.EqualTo(0));
        Assert.That(reader.GetNumberFormatIndex(1), Is.EqualTo(-1));
        Assert.That(reader.GetNumberFormatIndex(2), Is.EqualTo(14));
        Assert.That(reader.GetNumberFormatIndex(3), Is.EqualTo(164));
    }

    [Test]
    public void Issue11516_WorkbookWithSingleSheetShouldNotReturnEmptyDataset()
    {
        using IExcelDataReader reader = OpenReader("OldIssue11516_SingleTab");
        Assert.That(reader.ResultsCount, Is.EqualTo(1));

        DataSet dataSet = reader.AsDataSet();

        Assert.That(dataSet != null, Is.True);
        Assert.That(dataSet.Tables.Count, Is.EqualTo(1));
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(260));
        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(29));
    }

    [Test]
    public void Issue241_FirstOddEven()
    {
        using var reader = OpenReader("Issue224_Firstoddeven");
        Assert.That(reader.HeaderFooter, Is.Not.Null);

        Assert.That(reader.HeaderFooter?.HasDifferentFirst, Is.True, "HasDifferentFirst");
        Assert.That(reader.HeaderFooter?.HasDifferentOddEven, Is.True, "HasDifferentOddEven");

        Assert.That(reader.HeaderFooter?.FirstHeader, Is.EqualTo("&CFirst header center"), "First Header");
        Assert.That(reader.HeaderFooter?.FirstFooter, Is.EqualTo("&CFirst footer center"), "First Footer");
        Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft åäö &T&COdd page header&RRight  åäö &P"), "Odd Header");
        Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LLeft åäö &P&COdd Footer åäö &P&RRight åäö &D"), "Odd Footer");
        Assert.That(reader.HeaderFooter?.EvenHeader, Is.EqualTo("&L&A&CEven page header"), "Even Header");
        Assert.That(reader.HeaderFooter?.EvenFooter, Is.EqualTo("&CEven page footer"), "Even Footer");
    }

    [TestCase("standard_AES128_SHA1_ECB_pwd_password", Description = "OpenXml standard encryption aes128+sha1")]
    [TestCase("standard_AES192_SHA1_ECB_pwd_password", Description = "OpenXml standard encryption aes192+sha1")]
    [TestCase("standard_AES256_SHA1_ECB_pwd_password", Description = "OpenXml standard encryption aes256+sha1")]
    public void Issue242_StandardEncryption(string file)
    {
        using var reader = OpenReader(
            OpenStream(file),
            new ExcelReaderConfiguration() { Password = "password" });
        reader.Read();
        Assert.That(reader.GetString(0), Is.EqualTo("Password: password"));
    }

    [TestCase("agile_AES128_MD5_CBC_pwd_password")]
    [TestCase("agile_AES128_SHA1_CBC_pwd_password")]
    [TestCase("agile_AES128_SHA384_CBC_pwd_password")]
    [TestCase("agile_AES128_SHA512_CBC_pwd_password")]
    [TestCase("agile_AES192_SHA512_CBC_pwd_password")]
    [TestCase("agile_AES256_SHA512_CBC_pwd_password")]
    [TestCase("agile_DESede_SHA384_CBC_pwd_password")]
    [TestCase("agile_DES_MD5_CBC_pwd_password")]
    [TestCase("agile_RC2_SHA1_CBC_pwd_password")]
    public void Issue242_AgileEncryption(string file)
    {
        // OpenXml agile encryption aes128+md5+cbc
        using var reader = OpenReader(
            OpenStream(file),
            new ExcelReaderConfiguration() { Password = "password" });
        reader.Read();
        Assert.That(reader.GetString(0), Is.EqualTo("Password: password"));
    }

    [Test]
    public void AgileEncryptionReadsAcrossPackageSegments()
    {
        using var reader = OpenReader(
            OpenStream("agile_AES128_SHA1_CBC_pwd_password"),
            new ExcelReaderConfiguration() { Password = "password" });

        // The encrypted package spans multiple 4096-byte segments; reading the dataset traverses its contents.
        var dataSet = reader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo("Password: password"));
    }

    [Test]
    public void OpenXmlThrowsInvalidPasswordForWrongPassword()
    {
        Assert.Throws<Exceptions.InvalidPasswordException>(() =>
        {
            using var reader = OpenReader(
                OpenStream("agile_AES128_MD5_CBC_pwd_password"),
                new ExcelReaderConfiguration() { Password = "wrongpassword" });
            reader.Read();
        });
    }

    [Test]
    public void OpenXmlThrowsInvalidPasswordNoPassword()
    {
        Assert.Throws<Exceptions.InvalidPasswordException>(() =>
        {
            using var reader = OpenReader("agile_AES128_MD5_CBC_pwd_password");
            reader.Read();
        });
    }

    [Test]
    public void OpenXmlThrowsEmptyZipFile()
    {
        Assert.Throws<Exceptions.HeaderException>(() =>
        {
            using var reader = OpenReader("EmptyZipFile");
            reader.Read();
        });
    }

    [TestCase("10x10", null, Description = "Verify the file stream is closed and disposed by the reader")]
    [TestCase("standard_AES128_SHA1_ECB_pwd_password", "password", Description = "Verify streams used by standard encryption are closed")]
    [TestCase("agile_AES128_MD5_CBC_pwd_password", "password", Description = "Verify streams used by agile encryption are closed")]
    public void Issue265_OpenXmlDisposed(string file, string password)
    {
        var stream = OpenStream(file);

        using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(
            stream,
            new ExcelReaderConfiguration() { Password = password }))
        {
            _ = excelReader.AsDataSet();
        }

        Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
    }

    [Test]
    public void Issue_341_Indent()
    {
        int[][] expected =
        {
            new[] { 2, 0, 0 },
            new[] { 2, 0, 0 },
            new[] { 3, 3, 4 },
            new[] { 1, 1, 1 }, // Merged cell
            new[] { 2, 0, 0 },
        };

        int index = 0;
        using var reader = OpenReader("Issue341");
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
    public void Issue_341_HorizontalAlignment()
    {
        HorizontalAlignment[][] expected =
        {
            new[] { HorizontalAlignment.Left, HorizontalAlignment.General, HorizontalAlignment.General },
            new[] { HorizontalAlignment.Distributed, HorizontalAlignment.General, HorizontalAlignment.General },
            new[] { HorizontalAlignment.Left, HorizontalAlignment.Left, HorizontalAlignment.Left },
            new[] { HorizontalAlignment.Left, HorizontalAlignment.Left, HorizontalAlignment.Left }, // Merged cell
            new[] { HorizontalAlignment.Left, HorizontalAlignment.General, HorizontalAlignment.General },
        };

        int index = 0;
        using var reader = OpenReader("Issue341");
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
    public void Issue682()
    {
        using var reader = OpenReader("Issue682");
        reader.Read();
        var values = new object[reader.FieldCount];
        reader.GetValues(values);
        Assert.That(values, Is.EqualTo(new object[] { 100, 10, 1 }));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void DateAndDurationFixturesMatchXlsAfterEarlyAndCompleteReset(bool singlePass)
    {
        using var expectedReader = ExcelReaderFactory.CreateBinaryReader(
            Configuration.GetTestWorkbook("Issue283_TimeSpan.xls"),
            new ExcelReaderConfiguration { SinglePassMode = singlePass });
        using var actualReader = OpenReader(
            OpenFixtureStream("Issue283_TimeSpan"),
            new ExcelReaderConfiguration { SinglePassMode = singlePass });

        Assert.That(expectedReader.Read(), Is.True);
        Assert.That(actualReader.Read(), Is.True);
        expectedReader.Reset();
        actualReader.Reset();

        for (int pass = 0; pass < 2; pass++)
        {
            int rows = 0;
            while (expectedReader.Read())
            {
                if (actualReader.Read())
                {
                    for (int column = 0; column < 3; column++)
                    {
                        Assert.That(actualReader.GetValue(column), Is.EqualTo(expectedReader.GetValue(column)));
                        Assert.That(actualReader.GetFieldType(column), Is.EqualTo(expectedReader.GetFieldType(column)));
                        Assert.That(actualReader.GetCellError(column), Is.EqualTo(expectedReader.GetCellError(column)));
                    }
                }
                else
                {
                    // Excel exports can retain different numbers of trailing styled rows.
                    Assert.That(rows, Is.GreaterThanOrEqualTo(9));
                    for (int column = 0; column < 3; column++)
                    {
                        Assert.That(expectedReader.IsDBNull(column), Is.True);
                        Assert.That(expectedReader.GetCellError(column), Is.Null);
                    }
                }

                rows++;
            }

            Assert.That(rows, Is.GreaterThanOrEqualTo(9));

            // Open XML fixtures retain trailing styled rows in single-pass mode.
            while (actualReader.Read())
            {
                for (int column = 0; column < 3; column++)
                    Assert.That(actualReader.IsDBNull(column), Is.True);
            }

            Assert.That(actualReader.NextResult(), Is.EqualTo(expectedReader.NextResult()));
            expectedReader.Reset();
            actualReader.Reset();
        }
    }

    [TestCase(false)]
    [TestCase(true)]
    public void OpenXmlIndexedReadingMergesWorksheetWideOverlaps(bool distantRevisit)
    {
        int[] order = distantRevisit
            ? [.. Enumerable.Range(0, 70), 0, 1, 75]
            : [0, 3, 1, 4, 2, 5, 7, 6, 8];
        var rows = order.Select((index, ordinal) => new WorksheetConsumptionTests.OrderingRow(index, ordinal % 2, ordinal + 0.25, 20 + ordinal, ordinal % 7 == 0)).ToArray();
        rows = [.. rows, new WorksheetConsumptionTests.OrderingRow(order.Max() + 3, 0, null, 42, false)];
        using var stream = new MemoryStream(CreateOrderingWorkbook(rows));
        using (var reader = OpenReader(stream, new ExcelReaderConfiguration { LeaveOpen = true }))
        {
            Assert.That(reader.RowCount, Is.EqualTo(rows.Max(row => row.Index) + 1));
            Assert.That(reader.FieldCount, Is.EqualTo(2));
            Assert.That(reader.Read(), Is.True);
            reader.Reset();
            for (int pass = 0; pass < 2; pass++)
            {
                for (int index = 0; index < reader.RowCount; index++)
                {
                    Assert.That(reader.Read(), Is.True, $"row {index}");
                    var fragments = rows.Where(row => row.Index == index).ToArray();
                    double height = fragments.Length == 0 ? 15 : fragments.Last().Hidden ? 0 : fragments.Last().Height;
                    Assert.That(reader.RowHeight, Is.EqualTo(height), $"row {index}");
                    for (int column = 0; column < 2; column++)
                    {
                        var writes = fragments.Where(row => row.Column == column && row.Value.HasValue).ToArray();
                        Assert.That(reader.GetValue(column), Is.EqualTo(writes.Length == 0 ? DBNull.Value : (object)writes.Last().Value!.Value), $"row {index}, column {column}");
                    }
                }

                Assert.That(reader.Read(), Is.False);
                Assert.That(reader.Read(), Is.False);
                reader.Reset();
            }
        }

        Assert.That(stream.CanRead, Is.True);
    }

    [Test]
    public void OpenXmlSinglePassMergesCurrentFragmentsButRejectsCompletedRowRevisits()
    {
        WorksheetConsumptionTests.OrderingRow[] rows =
        [
            new(0, 0, 1, 24, false),
            new(0, 1, 2, 9, false),
            new(40, 0, 3, 15, false),
            new(0, 0, 4, 15, false),
        ];
        using var reader = OpenReader(new MemoryStream(CreateOrderingWorkbook(rows)), new ExcelReaderConfiguration { SinglePassMode = true });
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(0), Is.EqualTo(1));
            Assert.That(reader.GetDouble(1), Is.EqualTo(2));
            Assert.That(reader.RowHeight, Is.EqualTo(9));
            for (int index = 1; index <= 40; index++)
                Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(0), Is.EqualTo(3));
            Assert.That(Assert.Throws<InvalidOperationException>(() => reader.Read()).Message, Does.Contain("SinglePassMode = false"));
            reader.Reset();
        }
    }

    [Test]
    public void OpenXmlReorderedLeadingGapsRetainHeights()
    {
        WorksheetConsumptionTests.OrderingRow[] rows = [new(4, 0, 4, 24, false), new(2, 0, 2, 18, false)];
        using var reader = OpenReader(new MemoryStream(CreateOrderingWorkbook(rows)));
        double[] heights = [0, 15, 18, 15, 24];
        for (int index = 0; index < heights.Length; index++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.RowHeight, Is.EqualTo(heights[index]));
            if (index is 2 or 4)
                Assert.That(reader.GetDouble(0), Is.EqualTo(index));
            else
                Assert.That(reader.IsDBNull(0), Is.True);
        }

        Assert.That(reader.Read(), Is.False);
    }

    private byte[] CreateOrderingWorkbook(WorksheetConsumptionTests.OrderingRow[] rows) =>
        WorksheetConsumptionTests.OrderingWorkbook(UsesBinaryWorksheet ? "xlsb" : "xlsx", rows, UsesStrictWorksheet);
}