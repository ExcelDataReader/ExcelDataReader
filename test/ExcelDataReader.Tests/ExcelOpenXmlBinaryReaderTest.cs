using System;
using System.IO;

using NUnit.Framework;

namespace ExcelDataReader.Tests
{
    [TestFixture]
    public class ExcelOpenXmlBinaryReaderTest : ExcelOpenXmlReaderBase
    {
        /// <inheritdoc />
        protected override Stream OpenStream(string name)
        {
            return Configuration.GetTestWorkbook(name + ".xlsb");
        }

        /// <inheritdoc />
        protected override IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null)
        {
            return ExcelReaderFactory.CreateOpenXmlReader(stream, configuration);
        }

        /// <inheritdoc />
        protected override DateTime GitIssue82TodayDate => new(2013, 4, 19);

        [Test]
        public void GitIssue635()
        {
            using var reader = OpenReader("Test_git_issue_635");
            var dataSet = reader.AsDataSet();
            Assert.That(dataSet.Tables[0].Rows[0].ItemArray, Is.EqualTo(new[] { "A", "B", "C", "D", "E", "F" }));
        }
    }
}
