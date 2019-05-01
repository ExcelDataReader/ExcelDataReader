using System;
using System.IO;
using System.Text;
using ExcelDataReader.Tests;
using NUnit.Framework;

// Most CSV test data came from csv-spectrum: https://github.com/maxogden/csv-spectrum

#if EXCELDATAREADER_NET20
namespace ExcelDataReader.Net20.Tests
#elif NET45
namespace ExcelDataReader.Net45.Tests
#elif NETCOREAPP1_0
namespace ExcelDataReader.Netstandard13.Tests
#elif NETCOREAPP2_0
namespace ExcelDataReader.Netstandard20.Tests
#else
#error "Tests do not support the selected target platform"
#endif
{
    [TestFixture]

    public class ExcelCsvReaderTest
    {
        [OneTimeSetUp]
        public void TestInitialize()
        {
#if NETCOREAPP1_0 || NETCOREAPP2_0
            Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif
        }

        [Test]
        public void CsvCommaInQuotes()
        {
            using (var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("comma_in_quotes.csv")))
            {
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
        }

        [Test]
        public void CsvEscapedQuotes()
        {
            using (var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("escaped_quotes.csv")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual("a", ds.Tables[0].Rows[0][0]);
                Assert.AreEqual("b", ds.Tables[0].Rows[0][1]);
                Assert.AreEqual("1", ds.Tables[0].Rows[1][0]);
                Assert.AreEqual("ha \"ha\" ha", ds.Tables[0].Rows[1][1]);
                Assert.AreEqual("3", ds.Tables[0].Rows[2][0]);
                Assert.AreEqual("4", ds.Tables[0].Rows[2][1]);
            }
        }

        [Test]
        public void CsvQuotesAndNewlines()
        {
            using (var strm = new MemoryStream())
            {
                using (var writer = new StreamWriter(strm, Encoding.UTF8))
                {
                    writer.NewLine = "\n";
                    writer.WriteLine("a,b");
                    writer.WriteLine("1,\"ha ");
                    writer.WriteLine("\"\"ha\"\" ");
                    writer.WriteLine("ha\"");
                    writer.WriteLine("3,4");
                    writer.Flush();

                    using (var excelReader = ExcelReaderFactory.CreateCsvReader(strm))
                    {
                        var ds = excelReader.AsDataSet();
                        Assert.AreEqual("a", ds.Tables[0].Rows[0][0]);
                        Assert.AreEqual("b", ds.Tables[0].Rows[0][1]);

                        Assert.AreEqual("1", ds.Tables[0].Rows[1][0]);
                        Assert.AreEqual("ha \n\"ha\" \nha", ds.Tables[0].Rows[1][1]);

                        Assert.AreEqual("3", ds.Tables[0].Rows[2][0]);
                        Assert.AreEqual("4", ds.Tables[0].Rows[2][1]);
                    }
                }
            }
        }

        [Test]
        public void CsvEmpty()
        {
            // empty.csv
            // empty_crlf.csv
            TestEmpty("\n");
            TestEmpty("\r\n");
        }

        void TestEmpty(string linebreak)
        {
            using (var strm = new MemoryStream())
            {
                using (var writer = new StreamWriter(strm, Encoding.UTF8))
                {
                    writer.NewLine = linebreak;
                    writer.WriteLine("a,b,c");
                    writer.WriteLine("1,\"\",\"\"");
                    writer.WriteLine("2,3,4");
                    writer.Flush();

                    using (var excelReader = ExcelReaderFactory.CreateCsvReader(strm))
                    {
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
                }
            }
        }

        [Test]
        public void CsvNewlines()
        {
            // newlines.csv
            // newlines_crlf.csv
            TestNewlines("\n");
            TestNewlines("\r\n");
        }

        void TestNewlines(string linebreak)
        {
            using (var strm = new MemoryStream())
            {
                using (var writer = new StreamWriter(strm, Encoding.UTF8))
                {
                    writer.NewLine = linebreak;
                    writer.WriteLine("a,b,c");
                    writer.WriteLine("1,2,3");
                    writer.WriteLine("\"Once upon ");
                    writer.WriteLine("a time\",5,6");
                    writer.WriteLine("7,8,9");
                    writer.Flush();

                    using (var excelReader = ExcelReaderFactory.CreateCsvReader(strm))
                    {
                        var ds = excelReader.AsDataSet();
                        Assert.AreEqual("a", ds.Tables[0].Rows[0][0]);
                        Assert.AreEqual("b", ds.Tables[0].Rows[0][1]);
                        Assert.AreEqual("c", ds.Tables[0].Rows[0][2]);

                        Assert.AreEqual("1", ds.Tables[0].Rows[1][0]);
                        Assert.AreEqual("2", ds.Tables[0].Rows[1][1]);
                        Assert.AreEqual("3", ds.Tables[0].Rows[1][2]);

                        Assert.AreEqual("Once upon " + linebreak + "a time", ds.Tables[0].Rows[2][0]);
                        Assert.AreEqual("5", ds.Tables[0].Rows[2][1]);
                        Assert.AreEqual("6", ds.Tables[0].Rows[2][2]);

                        Assert.AreEqual("7", ds.Tables[0].Rows[3][0]);
                        Assert.AreEqual("8", ds.Tables[0].Rows[3][1]);
                        Assert.AreEqual("9", ds.Tables[0].Rows[3][2]);
                    }
                }
            }

        }

        [Test]
        public void CsvWhitespaceNull()
        {
            using (var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("simple_whitespace_null.csv")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual("a", ds.Tables[0].Rows[0][0]); // ignore spaces
                Assert.AreEqual("\0b\0", ds.Tables[0].Rows[0][1]);
                Assert.AreEqual("c", ds.Tables[0].Rows[0][2]); // ignore tabs

                Assert.AreEqual("1", ds.Tables[0].Rows[1][0]);
                Assert.AreEqual("2", ds.Tables[0].Rows[1][1]);
                Assert.AreEqual("3", ds.Tables[0].Rows[1][2]);
            }
        }

        [Test]
        public void CsvEncoding()
        {
            TestEncoding("utf8.csv", "ʤ");
            TestEncoding("utf8_bom.csv", "ʤ");
            TestEncoding("utf16le_bom.csv", "ʤ");
            TestEncoding("utf16be_bom.csv", "ʤ");
            TestEncoding("cp1252.csv", "æøå");
        }

        void TestEncoding(string workbook, string specialString)
        {
            using (var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook(workbook)))
            {
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

                using (var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("cp1252.csv"), configuration))
                {
                }
            });
        }

        [Test]
        public void CsvBigSheet()
        {
            using (var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("MOCK_DATA.csv")))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual("id", ds.Tables[0].Rows[0][0]);
                Assert.AreEqual("111.4.88.155", ds.Tables[0].Rows[1000][5]);

                // Check value at 1024 byte buffer boundary:
                // 17,Christoper,Blanning,cblanningg@so-net.ne.jp,Male,76.108.72.165
                Assert.AreEqual("cblanningg@so-net.ne.jp", ds.Tables[0].Rows[17][3]);

                Assert.AreEqual(6, ds.Tables[0].Columns.Count);
                Assert.AreEqual(1001, ds.Tables[0].Rows.Count);
            }
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
                AutodetectSeparators = new char[0]
            });

            TestNoSeparator(new ExcelReaderConfiguration()
            {
                AutodetectSeparators = null
            });
        }

        public void TestNoSeparator(ExcelReaderConfiguration configuration)
        {
            using (var strm = new MemoryStream())
            {
                using (var writer = new StreamWriter(strm, Encoding.UTF8))
                {
                    writer.WriteLine("This");
                    writer.WriteLine("is");
                    writer.WriteLine("a");
                    writer.Write("test");
                    writer.Flush();

                    using (var excelReader = ExcelReaderFactory.CreateCsvReader(strm, configuration))
                    {
                        var ds = excelReader.AsDataSet();
                        Assert.AreEqual("This", ds.Tables[0].Rows[0][0]);
                        Assert.AreEqual("is", ds.Tables[0].Rows[1][0]);
                        Assert.AreEqual("a", ds.Tables[0].Rows[2][0]);
                        Assert.AreEqual("test", ds.Tables[0].Rows[3][0]);
                    }
                }
            }
        }

        [Test]
        public void GitIssue_323_DoubleClose()
        {
            using (var reader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("MOCK_DATA.csv")))
            {
                reader.Read();
                reader.Close();
            }
        }

        [Test]
        public void GitIssue_333_EAN_Quotes()
        {
            using (var reader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("ean.txt")))
            {
                reader.Read();
                Assert.AreEqual(2, reader.RowCount);
                Assert.AreEqual(24, reader.FieldCount);
            }
        }

        [Test]
        public void GitIssue_351_Last_Line_Without_Line_Feed()
        {
            using (var strm = new MemoryStream())
            {
                using (var writer = new StreamWriter(strm, Encoding.UTF8))
                {
                    writer.NewLine = "\n";
                    writer.WriteLine("5;6;1;Test");
                    writer.Write("Test;;;");
                    writer.Flush();
                    using (var reader = ExcelReaderFactory.CreateCsvReader(strm))
                    {
                        var ds = reader.AsDataSet();
                        Assert.AreEqual(2, ds.Tables[0].Rows.Count);
                        Assert.AreEqual(4, reader.FieldCount);
                    }
                }
            }
        }

        [Test]
        public void ColumnWidthsTest()
        {
            using (var reader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("column_widths_test.csv")))
            {
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

                Assert.AreEqual($"Column at index 5 does not exist.{Environment.NewLine}Parameter name: i", 
                    exception.Message);
            }
        }

        [Test]
        public void CsvDisposed()
        {
            // Verify the file stream is closed and disposed by the reader
            {
                var stream = Configuration.GetTestWorkbook("MOCK_DATA.csv");
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateCsvReader(stream))
                {
                    var result = excelReader.AsDataSet();
                }

                Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
            }
        }

        [Test]
        public void CsvLeaveOpen()
        {
            // Verify the file stream is not disposed by the reader
            {
                var stream = Configuration.GetTestWorkbook("MOCK_DATA.csv");
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
                {
                    LeaveOpen = true
                }))
                {
                    var result = excelReader.AsDataSet();
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
                var stream = Configuration.GetTestWorkbook("MOCK_DATA.csv");
                using (IExcelDataReader reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
                {
                    AnalyzeInitialCsvRows = 100
                }))
                {
                    Assert.Throws(typeof(InvalidOperationException), () =>
                    {
                        var count = reader.RowCount;
                    });
                }
            }
        }
    }
}
