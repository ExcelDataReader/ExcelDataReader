using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
#if MSTEST_DEBUG || MSTEST_RELEASE
using Microsoft.VisualStudio.TestTools.UnitTesting;
#else
using NUnit.Framework;
#endif

namespace ExcelDataReader.Tests
{
    internal static class Configuration
    {
        private static readonly Dictionary<string, string> AppSettings = new Dictionary<string, string>
        {
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
            { "Test_Git_Issue_51", "Test_git_issue_51.xls" },
            { "Test_git_issue_111_NoRowRecords", "Test_git_issue_111_NoRowRecords.xls" },
            { "Test_git_Issue_142", "Test_git_Issue_142.xlsx" },
            { "Test_Git_Issue_145", "Test_git_issue_145.xls" },
            { "Test_git_issue_152", "Test_git_issue_152.xls" },
            { "Test_git_issue_158", "Test_git_issue_158.xls" },
            { "Test_git_issue_173", "Test_git_issue_173.xls" },
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
            { "xTest_Excel_OpenOffice", "Test_Excel_OpenOffice.xlsx" },
            { "Test_Issue_NoStyles_NoRAttribute", "Test_Issue_NoStyles_NoRAttribute.xlsx" },
            { "protectedsheet-xxx", "protectedsheet-xxx.xls" },
            { "TestTableOnlyImage_x01oct2016", "TestTableOnlyImage_x01oct2016.xls" },
            { "Test_InvalidByteOrderValueInHeader", "Test_InvalidByteOrderValueInHeader.xls" },
            { "AllColumnsNotReadInHiddenTable", "AllColumnsNotReadInHiddenTable.xls" },
            { "RowWithDifferentNumberOfColumns", "RowWithDifferentNumberOfColumns.xls" },
            { "NoDimensionOrCellReferenceAttribute", "NoDimensionOrCellReferenceAttribute.xlsx" },
            { "Test_Row1217NotRead", "Test_Row1217NotRead.xls" },
            { "StringContinuationAfterCharacterData", "StringContinuationAfterCharacterData.xls" },
            { "biff3", "biff3.xls" },
            { "Test_git_issue_5", "Test_git_issue_5.xls" },
            { "Test_git_issue_2", "Test_git_issue_2.xls" },
            { "ExcelLibrary_newdoc", "ExcelLibrary_newdoc.xls" },
            { "GitIssue_184_FATSectors", "GitIssue_184_FATSectors.xls" },
            { "Test_git_issue_217", "Test_git_issue_217.xls" },
            { "Test_git_issue_221", "Test_git_issue_221.xlsx" },
            { "Format49_@", "Format49_@.xlsx" },
            { "fillreport", "fillreport.xlsx" },
            { "Test_git_issue_231_NoCodePage", "Test_git_issue_231_NoCodePage.xls" },
            { "xroo_1900_base", "roo_1900_base.xlsx" },
            { "roo_1900_base", "roo_1900_base.xls" },
            { "xroo_1904_base", "roo_1904_base.xlsx" },
            { "roo_1904_base", "roo_1904_base.xls" },
            { "Test_git_issue_68_NullSheetPath", "Test_git_issue_68_NullSheetPath.xlsm" },
            { "Test_git_issue_53_Cached_Formula_String_Type", "Test_git_issue_53_Cached_Formula_String_Type.xlsx" },
            { "Test_git_issue_14_InvalidOADate", "Test_git_issue_14_InvalidOADate.xlsx" },
            { "as3xls_BIFF2", "as3xls_BIFF2.xls" },
            { "as3xls_BIFF3", "as3xls_BIFF3.xls" },
            { "as3xls_BIFF4", "as3xls_BIFF4.xls" },
            { "as3xls_BIFF5", "as3xls_BIFF5.xls" },
            { "Test_git_issue_224_simple_biff95", "Test_git_issue_224_simple_95.xls" },
            { "Test_git_issue_224_simple_biff", "Test_git_issue_224_simple.xls" },
            { "Test_git_issue_224_simple", "Test_git_issue_224_simple.xlsx" },
            { "Test_git_issue_224_firstoddeven", "Test_git_issue_224_firstoddeven.xlsx" },
            { "Test_git_issue_250_richtext", "Test_git_issue_250_richtext.xls" },
            { "xTest_git_issue_250_richtext", "Test_git_issue_250_richtext.xlsx" },
            { "Test_git_issue_242_std_rc4_pwd_password", "Test_git_issue_242_std_rc4_pwd_password.xls" },
            { "Test_git_issue_242_xor_pwd_password", "Test_git_issue_242_xor_pwd_password.xls" },
            { "standard_AES128_SHA1_ECB_pwd_password", "standard_AES128_SHA1_ECB_pwd_password.xlsx" },
            { "standard_AES192_SHA1_ECB_pwd_password", "standard_AES192_SHA1_ECB_pwd_password.xlsx" },
            { "standard_AES256_SHA1_ECB_pwd_password", "standard_AES256_SHA1_ECB_pwd_password.xlsx" },
            { "agile_AES128_MD5_CBC_pwd_password", "agile_AES128_MD5_CBC_pwd_password.xlsx" },
            { "agile_AES128_SHA1_CBC_pwd_password", "agile_AES128_SHA1_CBC_pwd_password.xlsx" },
            { "agile_AES128_SHA384_CBC_pwd_password", "agile_AES128_SHA384_CBC_pwd_password.xlsx" },
            { "agile_AES128_SHA512_CBC_pwd_password", "agile_AES128_SHA512_CBC_pwd_password.xlsx" },
            { "agile_AES192_SHA512_CBC_pwd_password", "agile_AES192_SHA512_CBC_pwd_password.xlsx" },
            { "agile_AES256_SHA512_CBC_pwd_password", "agile_AES256_SHA512_CBC_pwd_password.xlsx" },
            { "agile_DES_MD5_CBC_pwd_password", "agile_DES_MD5_CBC_pwd_password.xlsx" },
            { "agile_DESede_SHA384_CBC_pwd_password", "agile_DESede_SHA384_CBC_pwd_password.xlsx" },
            { "agile_RC2_SHA1_CBC_pwd_password", "agile_RC2_SHA1_CBC_pwd_password.xlsx" },
            { "EmptyZipFile", "EmptyZipFile.xlsx" },
            { "Test_git_issue_263", "Test_git_issue_263.xls" },
            { "CollapsedHide", "CollapsedHide.xls" },
            { "xCollapsedHide", "CollapsedHide.xlsx" },
            { "Test_git_issue_270", "Test_git_issue_270.xls" },
            { "xTest_git_issue_270", "Test_git_issue_270.xlsx" },
            { "Test_git_issue_271_InvalidDimension", "Test_git_issue_271_InvalidDimension.xlsx" },
            { "Test_git_issue_286_SST", "Test_git_issue_286_SST.xls" },
            { "Test_git_issue_283_TimeSpan", "Test_git_issue_283_TimeSpan.xls" },
            { "xTest_git_issue_283_TimeSpan", "Test_git_issue_283_TimeSpan.xlsx" },
            { "Test_git_issue289", "Test_git_issue289.xlsx" },
            { "Test_MergedCell_Binary", "Test_MergedCell.xls" },
            { "Test_MergedCell_OpenXml", "Test_MergedCell.xlsx" },
            { "Test_git_issue_301_IgnoreCase", "Test_git_issue_301_IgnoreCase.xlsx" },
            { "comma_in_quotes.csv", "csv/comma_in_quotes.csv" },
            { "escaped_quotes.csv", "csv/escaped_quotes.csv" },
            { "json.csv", "csv/json.csv" },
            { "simple_whitespace_null.csv", "csv/simple_whitespace_null.csv" },
            { "cp1252.csv", "csv/cp1252.csv" },
            { "utf8.csv", "csv/utf8.csv" },
            { "utf8_bom.csv", "csv/utf8_bom.csv" },
            { "utf16le_bom.csv", "csv/utf16le_bom.csv" },
            { "utf16be_bom.csv", "csv/utf16be_bom.csv" },
            { "MOCK_DATA.csv", "csv/MOCK_DATA.csv" },
            { "ean.txt", "csv/ean.txt" },
            { "Test_git_issue319", "Test_git_issue319.xlsx" },
            { "Test_git_issue321", "Test_git_issue321.xls" },
            { "Test_git_issue_324", "Test_git_issue_324.xlsx" },
            { "Test_git_issue_329_error.xlsx", "Test_git_issue_329_error.xlsx" },
            { "Test_git_issue_329_error.xls", "Test_git_issue_329_error.xls" },
            { "test_git_issue_354.xlsx", "test_git_issue_354.xlsx" },
            { "test_git_issue_364.xlsx", "test_git_issue_364.xlsx" },
            { "Test_git_issue_368_header.xls", "Test_git_issue_368_header.xls" },
            { "Test_git_issue_368_formats.xls", "Test_git_issue_368_formats.xls" },
            { "Test_git_issue_368_ixfe.xls", "Test_git_issue_368_ixfe.xls" },
            { "Test_git_issue_368_label_xf.xls", "Test_git_issue_368_label_xf.xls" },
            { "column_widths_test.csv", @"csv/column_widths_test.csv" },
            { "ColumnWidthsTest.xlsx", @"ColumnWidthsTest.xlsx" },
            { "ColumnWidthsTest.xls", @"ColumnWidthsTest.xls" },
            { "Test_git_issue_375_ixfe_rowmap.xls", "Test_git_issue_375_ixfe_rowmap.xls" },
            { "Test_git_issue_382_oom.xls", "Test_git_issue_382_oom.xls" },
            { "Test_git_issue_385_backslash.xlsx", "Test_git_issue_385_backslash.xlsx" },
            { "Test_git_issue_392_oob.xls", "Test_git_issue_392_oob.xls" },
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

        public static string GetTestWorkbookPath(string key)
        {
            var resources = Path.Combine(TestContext.CurrentContext.TestDirectory, "../../../../Resources");
            var fileName = Path.Combine(resources, GetKey(key));
            fileName = Path.GetFullPath(fileName);
            Assert.IsTrue(File.Exists(fileName), string.Format("By the key '{0}' the file '{1}' could not be found. Inside the Excel.Tests App.config file, edit the key basePath to be the folder where the test workbooks are located. If this is fine, check the filename that is related to the key.", key, fileName));
            return fileName;
        }

        public static ExcelDataSetConfiguration NoColumnNamesConfiguration = new ExcelDataSetConfiguration()
        {
            ConfigureDataTable = (reader) => new ExcelDataTableConfiguration()
            {
                UseHeaderRow = false
            }
        };

        public static ExcelDataSetConfiguration FirstRowColumnNamesConfiguration = new ExcelDataSetConfiguration()
        {
            ConfigureDataTable = (reader) => new ExcelDataTableConfiguration()
            {
                UseHeaderRow = true
            }
        };

        public static ExcelDataSetConfiguration FirstRowColumnNamesPrefixConfiguration = new ExcelDataSetConfiguration()
        {
            ConfigureDataTable = (reader) => new ExcelDataTableConfiguration()
            {
                UseHeaderRow = true,
                EmptyColumnNamePrefix = "Prefix"
            }
        };

    }
}