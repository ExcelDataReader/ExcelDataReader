using System;
using System.IO;
using NUnit.Framework;

namespace ExcelDataReader.Tests
{
    public class ExcelOpenXmlStrictReaderTest : ExcelOpenXmlReaderBase
    {
        protected override DateTime GitIssue82TodayDate => new(2013, 4, 19);

        protected override IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null)
        {
            return ExcelReaderFactory.CreateOpenXmlReader(stream, configuration);
        }
        protected override Stream OpenStream(string name)
        {
            return Configuration.GetTestWorkbook("strict\\" + name + ".xlsx");
        }

        [TestCase("Test_git_issue_498")]
        public void GitIssue498ReadStrictOpenXmlExcelFile(string fileName)
        {
            using IExcelDataReader reader = OpenReader(fileName);
            DataTableCollection tables = reader.AsDataSet().Tables;

            Assert.AreEqual(2, tables.Count);

            foreach (DataTable table in tables)
            {
                Assert.AreEqual(2, table.Rows.Count);
                Assert.AreEqual(2, table.Columns.Count);
                Assert.AreEqual("A1", table.Rows[0][0].ToString());
            }
        }
    }
}