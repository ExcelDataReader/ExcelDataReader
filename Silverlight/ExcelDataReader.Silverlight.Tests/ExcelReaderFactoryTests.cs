namespace ExcelDataReader.Silverlight.Tests
{
    using System;
    using System.IO;
    using System.Windows;
    using System.Windows.Resources;
    using Microsoft.Silverlight.Testing;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

	[TestClass]
	public class ExcelReaderFactoryTests
	{
		[TestMethod]
        [Tag("Binary")]
		public void CreateBinaryReader_WithFileStream_ReturnsValidBinaryReader()
		{
            StreamResourceInfo sri = Application.GetResourceStream(new Uri("Resources/BinaryFile.xls", UriKind.Relative));
            Stream fileStream = sri.Stream;

            var binaryReader = ExcelReaderFactory.CreateBinaryReader(fileStream);

            Assert.IsTrue(binaryReader.IsValid);
		}

        [TestMethod]
        [Tag("Xml")]
		public void CreateOpenXmlReader_WithFileStream_ReturnsValidOpenXmlReader()
		{
            StreamResourceInfo sri = Application.GetResourceStream(new Uri("Resources/OpenXmlFile.xlsx", UriKind.Relative));
            Stream fileStream = sri.Stream;

            var openXmlReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);

            Assert.IsTrue(openXmlReader.IsValid);
		}
	}
}