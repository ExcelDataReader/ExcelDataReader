using ExcelDataReader.Core;
using NUnit.Framework;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestMethod = NUnit.Framework.TestAttribute;

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
    [TestClass]
    public class FormatReaderTest
    {
        [TestMethod]
        public void Test_IsDateFormatString()
        {
            Assert.IsTrue(new FormatReader { FormatString = "dd/mm/yyyy" }.IsDateFormatString());
            Assert.IsTrue(new FormatReader { FormatString = "dd-mmm-yy" }.IsDateFormatString());
            Assert.IsTrue(new FormatReader { FormatString = "dd-mmmm" }.IsDateFormatString());
            Assert.IsTrue(new FormatReader { FormatString = "mmm-yy" }.IsDateFormatString());
            Assert.IsTrue(new FormatReader { FormatString = "h:mm AM/PM" }.IsDateFormatString());
            Assert.IsTrue(new FormatReader { FormatString = "h:mm:ss AM/PM" }.IsDateFormatString());
            Assert.IsTrue(new FormatReader { FormatString = "hh:mm" }.IsDateFormatString());
            Assert.IsTrue(new FormatReader { FormatString = "hh:mm:ss" }.IsDateFormatString());
            Assert.IsTrue(new FormatReader { FormatString = "dd/mm/yyyy hh:mm" }.IsDateFormatString());
            Assert.IsTrue(new FormatReader { FormatString = "mm:ss" }.IsDateFormatString());
            Assert.IsTrue(new FormatReader { FormatString = "mm:ss.0" }.IsDateFormatString());
            Assert.IsTrue(new FormatReader { FormatString = "[$-809]dd mmmm yyyy" }.IsDateFormatString());
            Assert.IsFalse(new FormatReader { FormatString = "#,##0;[Red]-#,##0" }.IsDateFormatString());
            Assert.IsFalse(new FormatReader { FormatString = "0_);[Red](0)" }.IsDateFormatString());
            Assert.IsFalse(new FormatReader { FormatString = @"0\h" }.IsDateFormatString());
            Assert.IsFalse(new FormatReader { FormatString = "0\"h\"" }.IsDateFormatString());
            Assert.IsFalse(new FormatReader { FormatString = "0%" }.IsDateFormatString());
            Assert.IsFalse(new FormatReader { FormatString = "General" }.IsDateFormatString());
            Assert.IsFalse(new FormatReader { FormatString = @"_-* #,##0\ _P_t_s_-;\-* #,##0\ _P_t_s_-;_-* "" - ""??\ _P_t_s_-;_-@_- " }.IsDateFormatString());
        }
    }
}
