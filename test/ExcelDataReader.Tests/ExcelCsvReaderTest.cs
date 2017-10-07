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
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
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
            using (var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook("quotes_and_newlines.csv")))
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

        [Test]
        public void CsvEmpty()
        {
            TestEmpty("empty.csv");
            TestEmpty("empty_crlf.csv");
        }

        void TestEmpty(string workbook)
        {
            using (var excelReader = ExcelReaderFactory.CreateCsvReader(Configuration.GetTestWorkbook(workbook)))
            {
                var ds = excelReader.AsDataSet();
                Assert.AreEqual("a", ds.Tables[0].Rows[0][0], workbook);
                Assert.AreEqual("b", ds.Tables[0].Rows[0][1], workbook);
                Assert.AreEqual("c", ds.Tables[0].Rows[0][2], workbook);

                Assert.AreEqual("1", ds.Tables[0].Rows[1][0], workbook);
                Assert.AreEqual("", ds.Tables[0].Rows[1][1], workbook);
                Assert.AreEqual("", ds.Tables[0].Rows[1][2], workbook);

                Assert.AreEqual("2", ds.Tables[0].Rows[2][0], workbook);
                Assert.AreEqual("3", ds.Tables[0].Rows[2][1], workbook);
                Assert.AreEqual("4", ds.Tables[0].Rows[2][2], workbook);
            }
        }

        [Test]
        public void CsvNewlines()
        {
            TestNewlines("newlines.csv", "\n");
            TestNewlines("newlines_crlf.csv", "\r\n");
        }

        void TestNewlines(string workbook, string linebreak)
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

                Assert.AreEqual("Once upon " + linebreak + "a time", ds.Tables[0].Rows[2][0], workbook);
                Assert.AreEqual("5", ds.Tables[0].Rows[2][1], workbook);
                Assert.AreEqual("6", ds.Tables[0].Rows[2][2], workbook);

                Assert.AreEqual("7", ds.Tables[0].Rows[3][0], workbook);
                Assert.AreEqual("8", ds.Tables[0].Rows[3][1], workbook);
                Assert.AreEqual("9", ds.Tables[0].Rows[3][2], workbook);
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
    }
}
