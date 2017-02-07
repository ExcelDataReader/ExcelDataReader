using System.Diagnostics;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
#if MSTEST_DEBUG || MSTEST_RELEASE
using Microsoft.VisualStudio.TestTools.UnitTesting;
#else
using NUnit.Framework;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestInitialize = NUnit.Framework.SetUpAttribute;
using TestCleanup = NUnit.Framework.TearDownAttribute;
using TestMethod = NUnit.Framework.TestAttribute;
#endif

namespace ExcelDataReader.Tests
{
    internal static class Configuration
    {
		public static Dictionary<string, string> AppSettings = new Dictionary<string, string> {
			{ "Test10x10", "Test10x10.xls" },
			{ "Test10x10000", "Test10x10000.xls" },
			{ "Test255x10", "Test255x10.xls" },
			{ "TestChess", "TestChess.xls" },
			{ "TestMultiSheet", "TestMultiSheet.xls" },
			{ "TestUnicodeChars", "TestUnicodeChars.xls" },
			{ "TestDoublePrecision", "TestDoublePrecision.xls" },
			{ "TestExcel_2040", "TestExcel_2040.xls" },
			{ "Test_Encoding_1520", "Test_Encoding_Formula_Date_1520.xls" },
			{ "Test_Decimal_1109", "Test_Decimal_1109.xls" },
			{ "Test_Issue_8536", "Test_Issue_8536.xls" },
			{ "Test_Issue_4031_NullColumn", "Test_Issue_4031_NullColumn.xls" },
			{ "Test_Issue_11397", "Test_Issue_11397.xls" },
			{ "Test_Issue_11435_Colors", "Test_Issue_11435_Colors.xls" },
			{ "Test_Issue_7433_IllegalOleAutDate", "Test_Issue_7433_IllegalOleAutDate.xls" },
			{ "Test_Issue_10725", "Test_Issue_10725.xls" },
			{ "Test_Issue_BoolFormula", "Test_Issue_BoolFormula.xls" },
			{ "TestFail_Binary", "TestFail_Binary.xls" },
			{ "Test_num_double_date_bool_string", "Test_num_double_date_bool_string.xls" },
			{ "Uncalculated", "Uncalculated.xls" },
			{ "Issue_11479_BlankSheet", "Test_Issue_11479_BlankSheet.xls" },
			{ "Test_BlankHeader", "Test_BlankHeader.xls" },
			{ "Test_OpenOffice", "Test_OpenOffice.xls" },
			{ "Test_Excel_Dataset", "Test_Excel_Dataset.xls" },
			{ "Test_Decimal_Locale", "Test_Decimal_Locale.xls" },
			{ "Test_Issue_11553_FAT", "Test_Issue_11553_FAT.xls" },
			{ "Test_Issue_11570_Excel2013", "Test_Issue_11570_Excel2013.xls" },
			{ "Test_Issue_11572_CodePage", "Test_Issue_11572_CodePage.xls" },
			{ "Test_Issue_11570_FAT_1", "Test_Issue_11570_FAT_1.xls" },
			{ "Test_Issue_11570_FAT_2", "Test_Issue_11570_FAT_2.xls" },
			{ "Test_Issue_11545_NoIndex", "Test_Issue_11545_NoIndex.xls" },
			{ "Test_Issue_11573_BlankValues", "Test_Issue_11573_BlankValues.xls" },
			{ "Test_Issue_DateFormatButNotDate", "Test_Issue_DateFormatButNotDate.xls" },
			{ "Test_Issue_11642_ValuesNotLoaded", "Test_Issue_11642_ValuesNotLoaded.xls" },
			{ "Test_Issue_11636_BiffStream", "Test_Issue_11636_BiffStream.xls" },
			{ "Test_Issue_12556_corrupt", "Test_Issue_12556_corrupt.xls" },
			{ "Test_Issue_11818_OutOfRange", "Test_Issue_11818_OutOfRange.xls" },
			{ "Test_Excel_OpenOffice", "Test_Excel_OpenOffice.xls" },
			{ "Test_Git_Issue_70", "Test_git_issue_70_ExcelBinaryReader_tryConvertOADateTime _convert_dates.xls" },
            { "Test_git_issue_45", "Test_git_issue_45.xls" },

			{ "xTestOpenXml", "TestOpenXml.xlsx" },
			{ "xTest10x10", "Test10x10.xlsx" },
			{ "xTest10x10000", "Test10x10000.xlsx" },
			{ "xTest255x10", "Test255x10.xlsx" },
			{ "xTestChess", "TestChess.xlsx" },
			{ "xTestMultiSheet", "TestMultiSheet.xlsx" },
			{ "xTestUnicodeChars", "TestUnicodeChars.xlsx" },
			{ "xTestDoublePrecision", "TestDoublePrecision.xlsx" },
			{ "xTestExcel_2040", "TestExcel_2040.xlsx" },
			{ "xTest_Encoding_1520", "Test_Encoding_Formula_Date_1520.xlsx" },
			{ "xTest_Decimal_1109", "Test_Decimal_1109.xlsx" },
			{ "xTest_Issue_8536", "Test_Issue_8536.xlsx" },
			{ "xTest_Issue_4031_NullColumn", "Test_Issue_4031_NullColumn.xlsx" },
			{ "xTest_Issue_11397", "Test_Issue_11397.xlsx" },
			{ "xTest_Issue_11435_Colors", "Test_Issue_11435_Colors.xlsx" },
			{ "xTest_Issue_7433_IllegalOleAutDate", "Test_Issue_7433_IllegalOleAutDate.xlsx" },
			{ "xTest_Issue_4145", "Test_Issue_4145.xlsx" },
			{ "xTest_Issue_10725", "Test_Issue_10725.xlsx" },
			{ "xTest_Issue_BoolFormula", "Test_Issue_BoolFormula.xlsx" },
			{ "xTestFail_Binary", "TestFail_Binary.xlsx" },
			{ "xTest_num_double_date_bool_string", "Test_num_double_date_bool_string.xlsx" },
			{ "xIssue_11479_BlankSheet", "Test_Issue_11479_BlankSheet.xlsx" },
			{ "xTest_BlankHeader", "Test_BlankHeader.xlsx" },
			{ "xTest_Excel_Dataset", "Test_Excel_Dataset.xlsx" },
			{ "xTest_Issue_11516_Single_Tab", "Test_Issue_11516_Single_Tab.xlsx" },
			{ "xTest_Decimal_Locale", "Test_Decimal_Locale.xlsx" },
			{ "xTest_Issue_xxx_LocaleTime", "Test_Issue_xxx_LocaleTime.xlsx" },
			{ "xTest_Issue_11522_OpenXml", "Test_Issue_11522_OpenXml.xlsx" },
			{ "xTest_Issue_11573_BlankValues", "Test_Issue_11573_BlankValues.xlsx" },
			{ "xTest_Issue_DateFormatButNotDate", "Test_Issue_DateFormatButNotDate.xlsx" },
			{ "xTest_Issue_11773_Exponential", "Test_Issue_11773_Exponential.xlsx" },
			{ "xTest_LotsOfSheets", "Test_LotsOfSheets.xlsx" },
			{ "bTest10x10", "Test10x10.xlsb" },
			{ "xTest_googlesourced", "Test_googlesourced.xlsx" },
			{ "xTest_Issue_12667_GoogleExport_MissingColumns", "Test_Issue_12667_GoogleExport_MissingColumns.xlsx" },
			{ "xTest_Excel_OpenOffice", "Test_Excel_OpenOffice.xlsx" }
		};

		public static Stream GetTestWorkbook(string key)
        {
            var fileName = GetTestWorkbookPath(key);
            return new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }

        public static string GetKey(string key)
        {
            string pathFile = /*ConfigurationManager.*/AppSettings[key];
            Debug.WriteLine(pathFile);
            return pathFile;
        }

        public static double ParseDouble(string s)
        {
            return double.Parse(s, CultureInfo.InvariantCulture);
        }

        public static string GetTestWorkbookPath(string key)
        {
			var fileName = Path.Combine(
				Path.Combine(TestContext.CurrentContext.WorkDirectory, "../Resources"), GetKey(key));
            //string fileName = Path.Combine(GetKey("basePath"), GetKey(key));
            fileName = Path.GetFullPath(fileName);
            Assert.IsTrue(File.Exists(fileName), string.Format("By the key '{0}' the file '{1}' could not be found. Inside the Excel.Tests App.config file, edit the key basePath to be the folder where the test workbooks are located. If this is fine, check the filename that is related to the key.", key, fileName));
            return fileName;
        }

    }
}