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
            Assert.AreEqual("first", ds.Tables[0].Rows[0][0]);
            Assert.AreEqual("last", ds.Tables[0].Rows[0][1]);
            Assert.AreEqual("address", ds.Tables[0].Rows[0][2]);
            Assert.AreEqual("city", ds.Tables[0].Rows[0][3]);
            Assert.AreEqual("zip", ds.Tables[0].Rows[0][4]);

            Assert.AreEqual("John", ds.Tables[0].Rows[1][0]);
            Assert.AreEqual("Doe", ds.Tables[0].Rows[1][1]);
            Assert.AreEqual("120 any st.", ds.Tables[0].Rows[1][2]);
            Assert.AreEqual("Anytown, WW", ds.Tables[0].Rows[1][3]);
            Assert.AreEqual("08123", ds.Tables[0].Rows[1][4]);
        }

        [Test]
        public void CsvEscapedQuotes()
        {
            using var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\escaped_quotes.csv"));
            var ds = excelReader.AsDataSet();
            Assert.AreEqual("a", ds.Tables[0].Rows[0][0]);
            Assert.AreEqual("b", ds.Tables[0].Rows[0][1]);
            Assert.AreEqual("1", ds.Tables[0].Rows[1][0]);
            Assert.AreEqual("ha \"ha\" ha", ds.Tables[0].Rows[1][1]);
            Assert.AreEqual("3", ds.Tables[0].Rows[2][0]);
            Assert.AreEqual("4", ds.Tables[0].Rows[2][1]);
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
            Assert.AreEqual("a", ds.Tables[0].Rows[0][0]);
            Assert.AreEqual("b", ds.Tables[0].Rows[0][1]);

            Assert.AreEqual("1", ds.Tables[0].Rows[1][0]);
            Assert.AreEqual("ha \n\"ha\" \nha", ds.Tables[0].Rows[1][1]);

            Assert.AreEqual("3", ds.Tables[0].Rows[2][0]);
            Assert.AreEqual("4", ds.Tables[0].Rows[2][1]);
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
            Assert.AreEqual("a", ds.Tables[0].Rows[0][0]);
            Assert.AreEqual("b", ds.Tables[0].Rows[0][1]);
            Assert.AreEqual("c", ds.Tables[0].Rows[0][2]);

            Assert.AreEqual("1", ds.Tables[0].Rows[1][0]);
            Assert.AreEqual("", ds.Tables[0].Rows[1][1]);
            Assert.AreEqual("", ds.Tables[0].Rows[1][2]);

            Assert.AreEqual("2", ds.Tables[0].Rows[2][0]);
            Assert.AreEqual("3", ds.Tables[0].Rows[2][1]);
            Assert.AreEqual("4", ds.Tables[0].Rows[2][2]);
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
            Assert.AreEqual("a", ds.Tables[0].Rows[0][0]);
            Assert.AreEqual("b", ds.Tables[0].Rows[0][1]);
            Assert.AreEqual("c", ds.Tables[0].Rows[0][2]);

            Assert.AreEqual("1", ds.Tables[0].Rows[1][0]);
            Assert.AreEqual("2", ds.Tables[0].Rows[1][1]);
            Assert.AreEqual("3", ds.Tables[0].Rows[1][2]);

            Assert.AreEqual("Once upon " + newLine + "a time", ds.Tables[0].Rows[2][0]);
            Assert.AreEqual("5", ds.Tables[0].Rows[2][1]);
            Assert.AreEqual("6", ds.Tables[0].Rows[2][2]);

            Assert.AreEqual("7", ds.Tables[0].Rows[3][0]);
            Assert.AreEqual("8", ds.Tables[0].Rows[3][1]);
            Assert.AreEqual("9", ds.Tables[0].Rows[3][2]);

        }

        [Test]
        public void CsvWhitespaceNull()
        {
            using var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\simple_whitespace_null.csv"));
            var ds = excelReader.AsDataSet();
            Assert.AreEqual("a", ds.Tables[0].Rows[0][0]); // ignore spaces
            Assert.AreEqual("\0b\0", ds.Tables[0].Rows[0][1]);
            Assert.AreEqual("c", ds.Tables[0].Rows[0][2]); // ignore tabs

            Assert.AreEqual("1", ds.Tables[0].Rows[1][0]);
            Assert.AreEqual("2", ds.Tables[0].Rows[1][1]);
            Assert.AreEqual("3", ds.Tables[0].Rows[1][2]);
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
            Assert.AreEqual("a", ds.Tables[0].Rows[0][0], workbook);
            Assert.AreEqual("b", ds.Tables[0].Rows[0][1], workbook);
            Assert.AreEqual("c", ds.Tables[0].Rows[0][2], workbook);
            Assert.AreEqual("1", ds.Tables[0].Rows[1][0], workbook);
            Assert.AreEqual("2", ds.Tables[0].Rows[1][1], workbook);
            Assert.AreEqual("3", ds.Tables[0].Rows[1][2], workbook);
            Assert.AreEqual("4", ds.Tables[0].Rows[2][0], workbook);
            Assert.AreEqual("5", ds.Tables[0].Rows[2][1], workbook);
            Assert.AreEqual(specialString, ds.Tables[0].Rows[2][2], workbook);
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
            Assert.AreEqual("id", ds.Tables[0].Rows[0][0]);
            Assert.AreEqual("111.4.88.155", ds.Tables[0].Rows[1000][5]);

            // Check value at 1024 byte buffer boundary:
            // 17,Christoper,Blanning,cblanningg@so-net.ne.jp,Male,76.108.72.165
            Assert.AreEqual("cblanningg@so-net.ne.jp", ds.Tables[0].Rows[17][3]);

            Assert.AreEqual(6, ds.Tables[0].Columns.Count);
            Assert.AreEqual(1001, ds.Tables[0].Rows.Count);
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
            Assert.AreEqual("This", ds.Tables[0].Rows[0][0]);
            Assert.AreEqual("is", ds.Tables[0].Rows[1][0]);
            Assert.AreEqual("a", ds.Tables[0].Rows[2][0]);
            Assert.AreEqual("test", ds.Tables[0].Rows[3][0]);
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
            Assert.AreEqual(2, reader.RowCount);
            Assert.AreEqual(24, reader.FieldCount);
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
            Assert.AreEqual(2, ds.Tables[0].Rows.Count);
            Assert.AreEqual(4, reader.FieldCount);
        }

        [Test]
        public void ColumnWidthsTest()
        {
            using var reader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("csv\\column_widths_test.csv"));
            reader.Read();
            Assert.AreEqual(8.43, reader.GetColumnWidth(0));
            Assert.AreEqual(8.43, reader.GetColumnWidth(1));
            Assert.AreEqual(8.43, reader.GetColumnWidth(2));
            Assert.AreEqual(8.43, reader.GetColumnWidth(3));
            Assert.AreEqual(8.43, reader.GetColumnWidth(4));

            var expectedException = typeof(ArgumentException);

            var exception = Assert.Throws(expectedException, () =>
            {
                reader.GetColumnWidth(5);
            });

#if NET5_0_OR_GREATER
                Assert.AreEqual($"Column at index 5 does not exist. (Parameter 'i')", 
                    exception.Message);
#else
            Assert.AreEqual($"Column at index 5 does not exist.{Environment.NewLine}Parameter name: i",
                exception.Message);
#endif
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
