using System.Text;
using ExcelDataReader.Desktop.Portable;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Excel.Tests
{
    [TestClass]
    public class FileHelperTest
    {
        [TestMethod]
        public void TestGetTempPath()
        {
            var fileHelper = new FileHelper();

            var tempPath = fileHelper.GetTempPath();

            tempPath.Should().NotBeNullOrEmpty();
        }
    }

    [TestClass]
    public class StringHelper
    {
        [TestMethod]
        public void TestIsSingleByte()
        {
            Encoding.UTF8.IsSingleByte.Should().BeFalse();
            Encoding.UTF7.IsSingleByte.Should().BeFalse();
            Encoding.UTF32.IsSingleByte.Should().BeFalse();
            Encoding.Unicode.IsSingleByte.Should().BeFalse();
            Encoding.ASCII.IsSingleByte.Should().BeTrue();
        }
    }
}
