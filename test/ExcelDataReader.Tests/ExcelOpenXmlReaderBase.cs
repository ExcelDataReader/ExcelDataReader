using NUnit.Framework;

namespace ExcelDataReader.Tests
{
    public abstract class ExcelOpenXmlReaderBase : ExcelTestBase
    {
        [Test]
        public void GitIssue14InvalidOADate()
        {
            using var excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_14_InvalidOADate.xlsx"));
            var dataSet = excelReader.AsDataSet();

            // Test out of range double formatted as date returns double
            Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo(1000000000000D));
        }

        [Test]
        public void GitIssue364()
        {
            using var reader = OpenReader("test_git_issue_364");
            Assert.That(reader.RowCount, Is.EqualTo(1));
            reader.Read();

            Assert.That(reader.GetNumberFormatIndex(0), Is.EqualTo(0));
            Assert.That(reader.GetNumberFormatIndex(1), Is.EqualTo(-1));
            Assert.That(reader.GetNumberFormatIndex(2), Is.EqualTo(14));
            Assert.That(reader.GetNumberFormatIndex(3), Is.EqualTo(164));
        }

        [Test]
        public void Issue11516WorkbookWithSingleSheetShouldNotReturnEmptyDataset()
        {
            using IExcelDataReader reader = OpenReader("Test_Issue_11516_Single_Tab");
            Assert.That(reader.ResultsCount, Is.EqualTo(1));

            DataSet dataSet = reader.AsDataSet();

            Assert.That(dataSet != null, Is.True);
            Assert.That(dataSet.Tables.Count, Is.EqualTo(1));
            Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(260));
            Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(29));
        }

        [Test]
        public void GitIssue241FirstOddEven()
        {
            using var reader = OpenReader("Test_git_issue_224_firstoddeven");
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


        // OpenXml standard encryption aes128+sha1
        [TestCase("standard_AES128_SHA1_ECB_pwd_password")]
        // OpenXml standard encryption aes192+sha1
        [TestCase("standard_AES192_SHA1_ECB_pwd_password")]
        // OpenXml standard encryption aes256+sha1
        [TestCase("standard_AES256_SHA1_ECB_pwd_password")]
        public void GitIssue242StandardEncryption(string file)
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
        public void GitIssue242AgileEncryption(string file)
        {
            // OpenXml agile encryption aes128+md5+cbc
            using var reader = OpenReader(
                OpenStream(file),
                new ExcelReaderConfiguration() { Password = "password" });
            reader.Read();
            Assert.That(reader.GetString(0), Is.EqualTo("Password: password"));
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

        // Verify the file stream is closed and disposed by the reader
        [TestCase("Test10x10", null)]
        // Verify streams used by standard encryption are closed
        [TestCase("standard_AES128_SHA1_ECB_pwd_password", "password")]
        // Verify streams used by agile encryption are closed
        [TestCase("agile_AES128_MD5_CBC_pwd_password", "password")]
        public void GitIssue265OpenXmlDisposed(string file, string password)
        {
            var stream = OpenStream(file);

            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(
                stream,
                new ExcelReaderConfiguration() { Password = password }))
            {
                var _ = excelReader.AsDataSet();
            }

            Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
        }

        [Test]
        public void GitIssue_341_Indent()
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
                new[] { HorizontalAlignment.Left, HorizontalAlignment.Left, HorizontalAlignment.Left }, // Merged cell
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
        public void GitIssue578()
        {

            using var reader = OpenReader("Test_git_issue578");

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
    }
}
