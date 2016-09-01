using System;
#if !LEGACY
using ExcelDataReader.Portable.Core;
#else
using Excel.Core;
#endif
using Microsoft.VisualStudio.TestTools.UnitTesting;

#if LEGACY
namespace Excel.Tests
#else
namespace ExcelDataReader.Tests
#endif
{
	[TestClass]
	public class FormatReaderTest
	{
		[TestMethod]
		public void Test_IsDateFormatString()
		{
	
			Assert.IsTrue(new FormatReader{FormatString = "dd/mm/yyyy"}.IsDateFormatString());
			Assert.IsTrue(new FormatReader{FormatString = "dd-mmm-yy"}.IsDateFormatString());
			Assert.IsTrue(new FormatReader{FormatString = "dd-mmmm"}.IsDateFormatString());
			Assert.IsTrue(new FormatReader{FormatString = "mmm-yy"}.IsDateFormatString());
			Assert.IsTrue(new FormatReader{FormatString = "h:mm AM/PM"}.IsDateFormatString());
			Assert.IsTrue(new FormatReader{FormatString = "h:mm:ss AM/PM"}.IsDateFormatString());
			Assert.IsTrue(new FormatReader{FormatString = "hh:mm"}.IsDateFormatString());
			Assert.IsTrue(new FormatReader{FormatString = "hh:mm:ss"}.IsDateFormatString());
			Assert.IsTrue(new FormatReader{FormatString = "dd/mm/yyyy hh:mm"}.IsDateFormatString());
			Assert.IsTrue(new FormatReader{FormatString = "mm:ss"}.IsDateFormatString());
			Assert.IsTrue(new FormatReader{FormatString = "mm:ss.0"}.IsDateFormatString());
			Assert.IsTrue(new FormatReader{FormatString = "[$-809]dd mmmm yyyy" }.IsDateFormatString());
			Assert.IsFalse(new FormatReader{FormatString = "#,##0;[Red]-#,##0" }.IsDateFormatString());
			Assert.IsFalse(new FormatReader{FormatString = "0_);[Red](0)" }.IsDateFormatString());
			Assert.IsFalse(new FormatReader{FormatString = @"0\h" }.IsDateFormatString());
			Assert.IsFalse(new FormatReader{FormatString = "0\"h\"" }.IsDateFormatString());
			Assert.IsFalse(new FormatReader{FormatString = "0%" }.IsDateFormatString());
			Assert.IsFalse(new FormatReader{FormatString = "General" }.IsDateFormatString());
		}
	}
}
