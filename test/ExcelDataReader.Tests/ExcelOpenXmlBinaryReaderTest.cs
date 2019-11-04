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
        protected override DateTime GitIssue82TodayDate => new DateTime(2013, 4, 19);
    }
}
