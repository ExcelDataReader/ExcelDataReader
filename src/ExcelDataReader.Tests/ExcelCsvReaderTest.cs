using System.Text;

namespace ExcelDataReader.Tests;

/// <summary>
/// Most CSV test data came from csv-spectrum: https://github.com/maxogden/csv-spectrum.
/// </summary>
public class ExcelCsvReaderTest
{
    [Test]
    public void CsvCommaInQuotes()
    {
        using var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\comma_in_quotes.csv"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds.Tables[0].Rows[0][0], Is.EqualTo("first"));
        Assert.That(ds.Tables[0].Rows[0][1], Is.EqualTo("last"));
        Assert.That(ds.Tables[0].Rows[0][2], Is.EqualTo("address"));
        Assert.That(ds.Tables[0].Rows[0][3], Is.EqualTo("city"));
        Assert.That(ds.Tables[0].Rows[0][4], Is.EqualTo("zip"));

        Assert.That(ds.Tables[0].Rows[1][0], Is.EqualTo("John"));
        Assert.That(ds.Tables[0].Rows[1][1], Is.EqualTo("Doe"));
        Assert.That(ds.Tables[0].Rows[1][2], Is.EqualTo("120 any st."));
        Assert.That(ds.Tables[0].Rows[1][3], Is.EqualTo("Anytown, WW"));
        Assert.That(ds.Tables[0].Rows[1][4], Is.EqualTo("08123"));
    }

    [Test]
    public void CsvEscapedQuotes()
    {
        using var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\escaped_quotes.csv"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds.Tables[0].Rows[0][0], Is.EqualTo("a"));
        Assert.That(ds.Tables[0].Rows[0][1], Is.EqualTo("b"));
        Assert.That(ds.Tables[0].Rows[1][0], Is.EqualTo("1"));
        Assert.That(ds.Tables[0].Rows[1][1], Is.EqualTo("ha \"ha\" ha"));
        Assert.That(ds.Tables[0].Rows[2][0], Is.EqualTo("3"));
        Assert.That(ds.Tables[0].Rows[2][1], Is.EqualTo("4"));
    }

    [Test]
    public void CsvQuotesAndNewlines()
    {
        using var stream = new MemoryStream();
        using var writer = new StreamWriter(stream, Encoding.UTF8);
        writer.NewLine = "\n";
        writer.WriteLine("a,b");
        writer.WriteLine("1,\"ha ");
        writer.WriteLine("\"\"ha\"\" ");
        writer.WriteLine("ha\"");
        writer.WriteLine("3,4");
        writer.Flush();

        using var excelReader = ExcelReaderFactory.CreateCsvReader(stream);
        var ds = excelReader.AsDataSet();
        Assert.That(ds.Tables[0].Rows[0][0], Is.EqualTo("a"));
        Assert.That(ds.Tables[0].Rows[0][1], Is.EqualTo("b"));

        Assert.That(ds.Tables[0].Rows[1][0], Is.EqualTo("1"));
        Assert.That(ds.Tables[0].Rows[1][1], Is.EqualTo("ha \n\"ha\" \nha"));

        Assert.That(ds.Tables[0].Rows[2][0], Is.EqualTo("3"));
        Assert.That(ds.Tables[0].Rows[2][1], Is.EqualTo("4"));
    }

    [Test]
    public void CsvEmpty()
    {
        // empty.csv
        // empty_crlf.csv
        TestEmpty("\n");
        TestEmpty("\r\n");

        static void TestEmpty(string newLine)
        {
            using var stream = new MemoryStream();
            using var writer = new StreamWriter(stream, Encoding.UTF8);
            writer.NewLine = newLine;
            writer.WriteLine("a,b,c");
            writer.WriteLine("1,\"\",\"\"");
            writer.WriteLine("2,3,4");
            writer.Flush();

            using var excelReader = ExcelReaderFactory.CreateCsvReader(stream);
            var ds = excelReader.AsDataSet();
            Assert.That(ds.Tables[0].Rows[0][0], Is.EqualTo("a"));
            Assert.That(ds.Tables[0].Rows[0][1], Is.EqualTo("b"));
            Assert.That(ds.Tables[0].Rows[0][2], Is.EqualTo("c"));

            Assert.That(ds.Tables[0].Rows[1][0], Is.EqualTo("1"));
            Assert.That(ds.Tables[0].Rows[1][1], Is.EqualTo(string.Empty));
            Assert.That(ds.Tables[0].Rows[1][2], Is.EqualTo(string.Empty));

            Assert.That(ds.Tables[0].Rows[2][0], Is.EqualTo("2"));
            Assert.That(ds.Tables[0].Rows[2][1], Is.EqualTo("3"));
            Assert.That(ds.Tables[0].Rows[2][2], Is.EqualTo("4"));
        }
    }

    [Test]
    public void CsvNewlines()
    {
        // newlines.csv
        // newlines_crlf.csv
        TestNewlines("\n");
        TestNewlines("\r\n");

        static void TestNewlines(string newLine)
        {
            using var stream = new MemoryStream();
            using var writer = new StreamWriter(stream, Encoding.UTF8);
            writer.NewLine = newLine;
            writer.WriteLine("a,b,c");
            writer.WriteLine("1,2,3");
            writer.WriteLine("\"Once upon ");
            writer.WriteLine("a time\",5,6");
            writer.WriteLine("7,8,9");
            writer.Flush();

            using var excelReader = ExcelReaderFactory.CreateCsvReader(stream);
            var ds = excelReader.AsDataSet();
            Assert.That(ds.Tables[0].Rows[0][0], Is.EqualTo("a"));
            Assert.That(ds.Tables[0].Rows[0][1], Is.EqualTo("b"));
            Assert.That(ds.Tables[0].Rows[0][2], Is.EqualTo("c"));

            Assert.That(ds.Tables[0].Rows[1][0], Is.EqualTo("1"));
            Assert.That(ds.Tables[0].Rows[1][1], Is.EqualTo("2"));
            Assert.That(ds.Tables[0].Rows[1][2], Is.EqualTo("3"));

            Assert.That(ds.Tables[0].Rows[2][0], Is.EqualTo("Once upon " + newLine + "a time"));
            Assert.That(ds.Tables[0].Rows[2][1], Is.EqualTo("5"));
            Assert.That(ds.Tables[0].Rows[2][2], Is.EqualTo("6"));

            Assert.That(ds.Tables[0].Rows[3][0], Is.EqualTo("7"));
            Assert.That(ds.Tables[0].Rows[3][1], Is.EqualTo("8"));
            Assert.That(ds.Tables[0].Rows[3][2], Is.EqualTo("9"));
        }
    }

    [Test]
    public void CsvWhitespaceNull()
    {
        using var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\simple_whitespace_null.csv"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds.Tables[0].Rows[0][0], Is.EqualTo("a")); // ignore spaces
        Assert.That(ds.Tables[0].Rows[0][1], Is.EqualTo("\0b\0"));
        Assert.That(ds.Tables[0].Rows[0][2], Is.EqualTo("c")); // ignore tabs

        Assert.That(ds.Tables[0].Rows[1][0], Is.EqualTo("1"));
        Assert.That(ds.Tables[0].Rows[1][1], Is.EqualTo("2"));
        Assert.That(ds.Tables[0].Rows[1][2], Is.EqualTo("3"));
    }

    [Test]
    public void CsvEncoding()
    {
        TestEncoding("csv\\utf8.csv", "ʤ");
        TestEncoding("csv\\utf8_bom.csv", "ʤ");
        TestEncoding("csv\\utf16le_bom.csv", "ʤ");
        TestEncoding("csv\\utf16be_bom.csv", "ʤ");
        TestEncoding("csv\\cp1252.csv", "æøå");

        static void TestEncoding(string workbook, string specialString)
        {
            using var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook(workbook));
            var ds = excelReader.AsDataSet();
            Assert.That(ds.Tables[0].Rows[0][0], Is.EqualTo("a"), workbook);
            Assert.That(ds.Tables[0].Rows[0][1], Is.EqualTo("b"), workbook);
            Assert.That(ds.Tables[0].Rows[0][2], Is.EqualTo("c"), workbook);
            Assert.That(ds.Tables[0].Rows[1][0], Is.EqualTo("1"), workbook);
            Assert.That(ds.Tables[0].Rows[1][1], Is.EqualTo("2"), workbook);
            Assert.That(ds.Tables[0].Rows[1][2], Is.EqualTo("3"), workbook);
            Assert.That(ds.Tables[0].Rows[2][0], Is.EqualTo("4"), workbook);
            Assert.That(ds.Tables[0].Rows[2][1], Is.EqualTo("5"), workbook);
            Assert.That(ds.Tables[0].Rows[2][2], Is.EqualTo(specialString), workbook);
        }
    }

    [Test]
    public void CsvWrongEncoding()
    {
        Assert.Throws(typeof(DecoderFallbackException), () =>
        {
            var configuration = new ExcelReaderConfiguration()
            {
                FallbackEncoding = Encoding.UTF8
            };

            using (ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\cp1252.csv"), configuration))
            {
            }
        });
    }

    [Test]
    public void CsvBigSheet()
    {
        using var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\MOCK_DATA.csv"));
        var ds = excelReader.AsDataSet();
        Assert.That(ds.Tables[0].Rows[0][0], Is.EqualTo("id"));
        Assert.That(ds.Tables[0].Rows[1000][5], Is.EqualTo("111.4.88.155"));

        // Check value at 1024 byte buffer boundary:
        // 17,Christoper,Blanning,cblanningg@so-net.ne.jp,Male,76.108.72.165
        Assert.That(ds.Tables[0].Rows[17][3], Is.EqualTo("cblanningg@so-net.ne.jp"));

        Assert.That(ds.Tables[0].Columns.Count, Is.EqualTo(6));
        Assert.That(ds.Tables[0].Rows.Count, Is.EqualTo(1001));
    }

    [Test]
    public void CsvNoSeparator()
    {
        TestNoSeparator(null);
    }

    [Test]
    public void CsvMissingSeparator()
    {
        TestNoSeparator(new ExcelReaderConfiguration()
        {
            AutodetectSeparators = Array.Empty<char>()
        });

        TestNoSeparator(new ExcelReaderConfiguration()
        {
            AutodetectSeparators = null
        });
    }

    public void TestNoSeparator(ExcelReaderConfiguration configuration)
    {
        using var stream = new MemoryStream();
        using var writer = new StreamWriter(stream, Encoding.UTF8);
        writer.WriteLine("This");
        writer.WriteLine("is");
        writer.WriteLine("a");
        writer.Write("test");
        writer.Flush();

        using var excelReader = ExcelReaderFactory.CreateCsvReader(stream, configuration);
        var ds = excelReader.AsDataSet();
        Assert.That(ds.Tables[0].Rows[0][0], Is.EqualTo("This"));
        Assert.That(ds.Tables[0].Rows[1][0], Is.EqualTo("is"));
        Assert.That(ds.Tables[0].Rows[2][0], Is.EqualTo("a"));
        Assert.That(ds.Tables[0].Rows[3][0], Is.EqualTo("test"));
    }

    [Test]
    public void GitIssue323DoubleClose()
    {
        using var reader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\MOCK_DATA.csv"));
        reader.Read();
        reader.Close();
    }

    [Test]
    public void GitIssue333EanQuotes()
    {
        using var reader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\ean.txt"));
        reader.Read();
        Assert.That(reader.RowCount, Is.EqualTo(2));
        Assert.That(reader.FieldCount, Is.EqualTo(24));
    }

    [Test]
    public void GitIssue351LastLineWithoutLineFeed()
    {
        using var stream = new MemoryStream();
        using var writer = new StreamWriter(stream, Encoding.UTF8);
        writer.NewLine = "\n";
        writer.WriteLine("5;6;1;Test");
        writer.Write("Test;;;");
        writer.Flush();
        using var reader = ExcelReaderFactory.CreateCsvReader(stream);
        var ds = reader.AsDataSet();
        Assert.That(ds.Tables[0].Rows.Count, Is.EqualTo(2));
        Assert.That(reader.FieldCount, Is.EqualTo(4));
    }

    [Test]
    public void ColumnWidthsTest()
    {
        using var reader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\column_widths_test.csv"));
        reader.Read();
        Assert.That(reader.GetColumnWidth(0), Is.EqualTo(8.43));
        Assert.That(reader.GetColumnWidth(1), Is.EqualTo(8.43));
        Assert.That(reader.GetColumnWidth(2), Is.EqualTo(8.43));
        Assert.That(reader.GetColumnWidth(3), Is.EqualTo(8.43));
        Assert.That(reader.GetColumnWidth(4), Is.EqualTo(8.43));

        var expectedException = typeof(ArgumentException);

        var exception = Assert.Throws(expectedException, () =>
        {
            reader.GetColumnWidth(5);
        });

        Assert.That(exception.Message, Does.StartWith($"Column at index 5 does not exist"));
    }

    [Test]
    public void CsvDisposed()
    {
        // Verify the file stream is closed and disposed by the reader
        {
            var stream = Configuration.GetTestWorkbook("csv\\MOCK_DATA.csv");
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateCsvReader(stream))
            {
                _ = excelReader.AsDataSet();
            }

            Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
        }
    }

    [Test]
    public void CsvLeaveOpen()
    {
        // Verify the file stream is not disposed by the reader
        {
            var stream = Configuration.GetTestWorkbook("csv\\MOCK_DATA.csv");
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
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
    public void CsvRowCountAnalyzeRowsThrows()
    {
        {
            var stream = Configuration.GetTestWorkbook("csv\\MOCK_DATA.csv");
            using IExcelDataReader reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
            {
                AnalyzeInitialCsvRows = 100
            });
            Assert.Throws(typeof(InvalidOperationException), () =>
            {
                _ = reader.RowCount;
            });
        }
    }

    [Test]
    public void GitIssue578()
    {
        var stream = Configuration.GetTestWorkbook(@"csv\Test_git_issue578.csv");
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration());
        excelReader.Read();
        var values = new object[excelReader.FieldCount];
        excelReader.GetValues(values);
        var values2 = new object[excelReader.FieldCount + 1];
        excelReader.GetValues(values2);
        var values3 = new object[excelReader.FieldCount - 1];
        excelReader.GetValues(values3);

        Assert.That(values, Is.EqualTo(new object[] { "1", "2", "3", "4", "5" }));
        Assert.That(values2, Is.EqualTo(new object[] { "1", "2", "3", "4", "5", null }));
        Assert.That(values3, Is.EqualTo(new object[] { "1", "2", "3", "4" }));
    }

    [Test]
    public void GitIssue500()
    {
        const string data = """
            Item_Number,Pmt_Amount,Type,Voided,Note
            200212812,$462.76,Check,06/06/2018,Hash#hash

            """;

        using var reader = ExcelReaderFactory.CreateCsvReader(new MemoryStream(Encoding.UTF8.GetBytes(data)));
        reader.Read();
        object[] row1 = new object[reader.FieldCount];
        reader.GetValues(row1);
        reader.Read();
        object[] row2 = new object[reader.FieldCount];
        reader.GetValues(row2);

        Assert.Multiple(() =>
        {
            Assert.That(row1, Is.EqualTo(new object[] { "Item_Number", "Pmt_Amount", "Type", "Voided", "Note" }));
            Assert.That(row2, Is.EqualTo(new object[] { "200212812", "$462.76", "Check", "06/06/2018", "Hash#hash" }));
        });
    }

    [Test]
    public void GitIssue500_QuotedValueWithNewLine()
    {
        const string data = """
            Item_Number,Pmt_Amount,"Type

            2",Voided,Note
            200212812,$462.76,Check,06/06/2018,Hash#hash

            """;

        using var reader = ExcelReaderFactory.CreateCsvReader(new MemoryStream(Encoding.UTF8.GetBytes(data)));
        reader.Read();
        object[] row1 = new object[reader.FieldCount];
        reader.GetValues(row1);
        reader.Read();
        object[] row2 = new object[reader.FieldCount];
        reader.GetValues(row2);

        Assert.Multiple(() =>
        {
            Assert.That(row1, Is.EqualTo(new object[]
            {
                 "Item_Number",
                 "Pmt_Amount", """
                 Type

                 2
                 """,
                 "Voided",
                 "Note",
            }));
            Assert.That(row2, Is.EqualTo(new object[] { "200212812", "$462.76", "Check", "06/06/2018", "Hash#hash" }));
        });
    }

    [Test]
    public void GitIssue463()
    {
#pragma warning disable SA1027 // TabsMustNotBeUsed
        const string data = """
            Name	Currency	Type	Cost	"Cost per 1,000 Items"
            Test1	ABC	XX	"10,143.27"	0.00
            Test2	EFG	YY	"10,143.27"	0.00
            Test3	IJK	ZZ	"10,143.27"	0.00

            """;
#pragma warning restore SA1027 // TabsMustNotBeUsed

        using var reader = ExcelReaderFactory.CreateCsvReader(new MemoryStream(Encoding.UTF8.GetBytes(data)));
        reader.Read();
        object[] row1 = new object[reader.FieldCount];
        reader.GetValues(row1);
        reader.Read();
        object[] row2 = new object[reader.FieldCount];
        reader.GetValues(row2);
        reader.Read();
        object[] row3 = new object[reader.FieldCount];
        reader.GetValues(row3);
        reader.Read();
        object[] row4 = new object[reader.FieldCount];
        reader.GetValues(row4);

        Assert.Multiple(() =>
        {
            Assert.That(row1, Is.EqualTo(new object[] { "Name", "Currency", "Type", "Cost", "Cost per 1,000 Items" }));
            Assert.That(row2, Is.EqualTo(new object[] { "Test1", "ABC", "XX", "10,143.27", "0.00" }));
            Assert.That(row3, Is.EqualTo(new object[] { "Test2", "EFG", "YY", "10,143.27", "0.00" }));
            Assert.That(row4, Is.EqualTo(new object[] { "Test3", "IJK", "ZZ", "10,143.27", "0.00" }));
        });
    }

    [Test]
    public void GitIssue642_ActiveSheet()
    {
        using var reader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\MOCK_DATA.csv"));
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            FilterSheet = (tableReader, sheetIndex) => tableReader.IsActiveSheet
        });
        Assert.That(reader.ActiveSheet, Is.EqualTo(0));
        Assert.That(dataSet.Tables.Count, Is.EqualTo(1));
    }

    [Test]
    public void GitIssue580_ReadCsvWithoutQuoteChar()
    {
#pragma warning disable SA1027 // TabsMustNotBeUsed
        const string data = """
            Name	Currency	Ty"pe	Cost	Cost
            Test1	A"BC	XX	10,143.27	"0.00
            Test2	EFG	YY	10,143.27	0.00
            Test3	IJK	ZZ"	10,143.27	0.00

            """;
#pragma warning restore SA1027 // TabsMustNotBeUsed

        using var reader = ExcelReaderFactory.CreateCsvReader(new MemoryStream(Encoding.UTF8.GetBytes(data)), new()
        {
            QuoteChar = null
        });
        reader.Read();
        object[] row1 = new object[reader.FieldCount];
        reader.GetValues(row1);
        reader.Read();
        object[] row2 = new object[reader.FieldCount];
        reader.GetValues(row2);
        reader.Read();
        object[] row3 = new object[reader.FieldCount];
        reader.GetValues(row3);
        reader.Read();
        object[] row4 = new object[reader.FieldCount];
        reader.GetValues(row4);

        Assert.Multiple(() =>
        {
            Assert.That(row1, Is.EqualTo(new object[] { "Name", "Currency", "Ty\"pe", "Cost", "Cost" }));
            Assert.That(row2, Is.EqualTo(new object[] { "Test1", "A\"BC", "XX", "10,143.27", "\"0.00" }));
            Assert.That(row3, Is.EqualTo(new object[] { "Test2", "EFG", "YY", "10,143.27", "0.00" }));
            Assert.That(row4, Is.EqualTo(new object[] { "Test3", "IJK", "ZZ\"", "10,143.27", "0.00" }));
        });
    }

    [Test]
    public void GitIssue580_ReadCsvWithCustomQuoteChar()
    {
#pragma warning disable SA1027 // TabsMustNotBeUsed
        const string data = """
            Name	Currency	Type	Cost	'Cost per 1,000 Items'
            Test1	ABC	XX	'10,143.27'	0.00
            Test2	EFG	YY	'10,143.27'	0.00
            Test3	IJK	ZZ	'10,143.27'	0.00

            """;
#pragma warning restore SA1027 // TabsMustNotBeUsed

        using var reader = ExcelReaderFactory.CreateCsvReader(new MemoryStream(Encoding.UTF8.GetBytes(data)), new()
        {
            QuoteChar = '\''
        });
        reader.Read();
        object[] row1 = new object[reader.FieldCount];
        reader.GetValues(row1);
        reader.Read();
        object[] row2 = new object[reader.FieldCount];
        reader.GetValues(row2);
        reader.Read();
        object[] row3 = new object[reader.FieldCount];
        reader.GetValues(row3);
        reader.Read();
        object[] row4 = new object[reader.FieldCount];
        reader.GetValues(row4);

        Assert.Multiple(() =>
        {
            Assert.That(row1, Is.EqualTo(new object[] { "Name", "Currency", "Type", "Cost", "Cost per 1,000 Items" }));
            Assert.That(row2, Is.EqualTo(new object[] { "Test1", "ABC", "XX", "10,143.27", "0.00" }));
            Assert.That(row3, Is.EqualTo(new object[] { "Test2", "EFG", "YY", "10,143.27", "0.00" }));
            Assert.That(row4, Is.EqualTo(new object[] { "Test3", "IJK", "ZZ", "10,143.27", "0.00" }));
        });
    }
}