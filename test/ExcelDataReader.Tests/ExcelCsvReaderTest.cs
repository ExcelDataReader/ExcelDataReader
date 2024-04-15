using System;
using System.IO;
using System.Text;
using NUnit.Framework;

// Most CSV test data came from csv-spectrum: https://github.com/maxogden/csv-spectrum

namespace ExcelDataReader.Tests
{
    [TestFixture]

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
        }

        private static void TestEmpty(string newLine)
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
            Assert.That(ds.Tables[0].Rows[1][1], Is.EqualTo(""));
            Assert.That(ds.Tables[0].Rows[1][2], Is.EqualTo(""));

            Assert.That(ds.Tables[0].Rows[2][0], Is.EqualTo("2"));
            Assert.That(ds.Tables[0].Rows[2][1], Is.EqualTo("3"));
            Assert.That(ds.Tables[0].Rows[2][2], Is.EqualTo("4"));
        }

        [Test]
        public void CsvNewlines()
        {
            // newlines.csv
            // newlines_crlf.csv
            TestNewlines("\n");
            TestNewlines("\r\n");
        }

        private static void TestNewlines(string newLine)
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
        }

        private static void TestEncoding(string workbook, string specialString)
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

        [Test]
        public void CsvWrongEncoding()
        {
            Assert.Throws(typeof(DecoderFallbackException), () =>
            {
                var configuration = new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.UTF8
                };

                using var _ = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\cp1252.csv"), configuration);
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
                    var _ = excelReader.AsDataSet();
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
                    var _ = excelReader.AsDataSet();
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
                    var _ = reader.RowCount;
                });
            }
        }
    }
}
