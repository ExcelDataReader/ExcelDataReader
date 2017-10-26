using NUnit.Framework;
using System;
using System.Globalization;
using TestClass = NUnit.Framework.TestFixtureAttribute;
using TestMethod = NUnit.Framework.TestAttribute;

// Much of the test data was adapted from the SheetJS/ssf project.
// Copyright (C) 2013-present  SheetJS. 
// Licensed under the Apache License, Version 2.0


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
        string Format(object value, string formatString, CultureInfo culture)
        {
            var format = new NumberFormatString(formatString);
            if (format.IsValid)
                return format.Format(value, culture);

            return null;
        }

        bool IsDateFormatString(string formatString)
        {
            var format = new NumberFormatString(formatString);
            return format?.IsDateTimeFormat ?? false;
        }

        [TestMethod]
        public void NumberFormat_TestFractionAlignmentSuffix()
        {
            Test(0, "??/??", " 0/1 ");
            Test(1.5, "??/??", " 3/2 ");
            Test(3.4, "??/??", "17/5 ");
            Test(4.3, "??/??", "43/10");

            Test(0, "00/00", "00/01");
            Test(1.5, "00/00", "03/02");
            Test(3.4, "00/00", "17/05");
            Test(4.3, "00/00", "43/10");

            Test(0.00, "# ??/\"a\"?\"a\"0\"a\"", "0        a");
            Test(0.10, "# ??/\"a\"?\"a\"0\"a\"", "0        a");
            Test(0.12, "# ??/\"a\"?\"a\"0\"a\"", "  1/a8a0a");

            Test(1.00, "# ??/\"a\"?\"a\"0\"a\"", "1        a");
            Test(1.10, "# ??/\"a\"?\"a\"0\"a\"", "1  1/a9a0a");
            Test(1.12, "# ??/\"a\"?\"a\"0\"a\"", "1  1/a8a0a");
        }


        [TestMethod]
        public void NumberFormat_TestIsDateFormatString()
        {
            Assert.IsTrue(IsDateFormatString("dd/mm/yyyy"));
            Assert.IsTrue(IsDateFormatString("dd-mmm-yy"));
            Assert.IsTrue(IsDateFormatString("dd-mmmm"));
            Assert.IsTrue(IsDateFormatString("mmm-yy"));
            Assert.IsTrue(IsDateFormatString("h:mm AM/PM"));
            Assert.IsTrue(IsDateFormatString("h:mm:ss AM/PM"));
            Assert.IsTrue(IsDateFormatString("hh:mm"));
            Assert.IsTrue(IsDateFormatString("hh:mm:ss"));
            Assert.IsTrue(IsDateFormatString("dd/mm/yyyy hh:mm"));
            Assert.IsTrue(IsDateFormatString("mm:ss"));
            Assert.IsTrue(IsDateFormatString("mm:ss.0"));
            Assert.IsTrue(IsDateFormatString("[$-809]dd mmmm yyyy"));
            Assert.IsFalse(IsDateFormatString("#,##0;[Red]-#,##0"));
            Assert.IsFalse(IsDateFormatString("0_);[Red](0)"));
            Assert.IsFalse(IsDateFormatString(@"0\h"));
            Assert.IsFalse(IsDateFormatString("0\"h\""));
            Assert.IsFalse(IsDateFormatString("0%"));
            Assert.IsFalse(IsDateFormatString("General"));
            Assert.IsFalse(IsDateFormatString(@"_-* #,##0\ _P_t_s_-;\-* #,##0\ _P_t_s_-;_-* "" - ""??\ _P_t_s_-;_-@_- "));
        }

        [TestMethod]
        public void NumberFormat_TestDateTime()
        {
            Test(new DateTime(2000, 1, 1), "d-mmm-yy", "1-Jan-00");
            Test(new DateTime(2000, 1, 1, 12, 34, 56), "m/d/yyyy\\ h:mm:ss;@", "1/1/2000 12:34:56");
        }

        [TestMethod]
        public void NumberFormat_TestTimeSpan()
        {
            Test(new TimeSpan(100, 0, 0), "[hh]:mm:ss", "100:00:00");
            Test(new TimeSpan(100, 0, 0), "[mm]:ss", "6000:00");
        }

        void Test(object value, string format, string expected)
        {
            var result = Format(value, format, CultureInfo.InvariantCulture);
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void NumberFormat_TestFraction()
        {
            Test(1, "# ?/?", "1    ");
            Test(-1.2, "# ?/?", "-1 1/5");
            Test(12.3, "# ?/?", "12 1/3");
            Test(-12.34, "# ?/?", "-12 1/3");
            Test(123.45, "# ?/?", "123 4/9");
            Test(-123.456, "# ?/?", "-123 1/2");
            Test(1234.567, "# ?/?", "1234 4/7");
            Test(-1234.5678, "# ?/?", "-1234 4/7");
            Test(12345.6789, "# ?/?", "12345 2/3");
            Test(-12345.67891, "# ?/?", "-12345 2/3");

            Test(1, "# ??/??", "1      ");
            Test(-1.2, "# ??/??", "-1  1/5 ");
            Test(12.3, "# ??/??", "12  3/10");
            Test(-12.34, "# ??/??", "-12 17/50");
            Test(123.45, "# ??/??", "123  9/20");
            Test(-123.456, "# ??/??", "-123 26/57");
            Test(1234.567, "# ??/??", "1234 55/97");
            Test(-1234.5678, "# ??/??", "-1234 46/81");
            Test(12345.6789, "# ??/??", "12345 55/81");
            Test(-12345.67891, "# ??/??", "-12345 55/81");

            Test(1, "# ???/???", "1        ");
            Test(-1.2, "# ???/???", "-1   1/5  ");
            Test(12.3, "# ???/???", "12   3/10 ");
            Test(-12.34, "# ???/???", "-12  17/50 ");
            Test(123.45, "# ???/???", "123   9/20 ");
            Test(-123.456, "# ???/???", "-123  57/125");
            Test(1234.567, "# ???/???", "1234  55/97 ");
            Test(-1234.5678, "# ???/???", "-1234  67/118");
            Test(12345.6789, "# ???/???", "12345  74/109");
            Test(-12345.67891, "# ???/???", "-12345 573/844");


            Test(1, "# ?/2", "1    ");
            Test(-1.2, "# ?/2", "-1    ");
            Test(12.3, "# ?/2", "12 1/2");
            Test(-12.34, "# ?/2", "-12 1/2");
            Test(123.45, "# ?/2", "123 1/2");
            Test(-123.456, "# ?/2", "-123 1/2");
            Test(1234.567, "# ?/2", "1234 1/2");
            Test(-1234.5678, "# ?/2", "-1234 1/2");
            Test(12345.6789, "# ?/2", "12345 1/2");
            Test(-12345.67891, "# ?/2", "-12345 1/2");

            Test(1, "# ?/4", "1    ");
            Test(-1.2, "# ?/4", "-1 1/4");
            Test(12.3, "# ?/4", "12 1/4");
            Test(-12.34, "# ?/4", "-12 1/4");
            Test(123.45, "# ?/4", "123 2/4");
            Test(-123.456, "# ?/4", "-123 2/4");
            Test(1234.567, "# ?/4", "1234 2/4");
            Test(-1234.5678, "# ?/4", "-1234 2/4");
            Test(12345.6789, "# ?/4", "12345 3/4");
            Test(-12345.67891, "# ?/4", "-12345 3/4");

            Test(1, "# ?/8", "1    ");
            Test(-1.2, "# ?/8", "-1 2/8");
            Test(12.3, "# ?/8", "12 2/8");
            Test(-12.34, "# ?/8", "-12 3/8");
            Test(123.45, "# ?/8", "123 4/8");
            Test(-123.456, "# ?/8", "-123 4/8");
            Test(1234.567, "# ?/8", "1234 5/8");
            Test(-1234.5678, "# ?/8", "-1234 5/8");
            Test(12345.6789, "# ?/8", "12345 5/8");
            Test(-12345.67891, "# ?/8", "-12345 5/8");

            Test(1, "# ??/16", "1      ");
            Test(-1.2, "# ??/16", "-1  3/16");
            Test(12.3, "# ??/16", "12  5/16");
            Test(-12.34, "# ??/16", "-12  5/16");
            Test(123.45, "# ??/16", "123  7/16");
            Test(-123.456, "# ??/16", "-123  7/16");
            Test(1234.567, "# ??/16", "1234  9/16");
            Test(-1234.5678, "# ??/16", "-1234  9/16");
            Test(12345.6789, "# ??/16", "12345 11/16");
            Test(-12345.67891, "# ??/16", "-12345 11/16");

            Test(1, "# ?/10", "1     ");
            Test(-1.2, "# ?/10", "-1 2/10");
            Test(12.3, "# ?/10", "12 3/10");
            Test(-12.34, "# ?/10", "-12 3/10");
            Test(123.45, "# ?/10", "123 5/10");
            Test(-123.456, "# ?/10", "-123 5/10");
            Test(1234.567, "# ?/10", "1234 6/10");
            Test(-1234.5678, "# ?/10", "-1234 6/10");
            Test(12345.6789, "# ?/10", "12345 7/10");
            Test(-12345.67891, "# ?/10", "-12345 7/10");

            Test(1, "# ??/100", "1       ");
            Test(-1.2, "# ??/100", "-1 20/100");
            Test(12.3, "# ??/100", "12 30/100");
            Test(-12.34, "# ??/100", "-12 34/100");
            Test(123.45, "# ??/100", "123 45/100");
            Test(-123.456, "# ??/100", "-123 46/100");
            Test(1234.567, "# ??/100", "1234 57/100");
            Test(-1234.5678, "# ??/100", "-1234 57/100");
            Test(12345.6789, "# ??/100", "12345 68/100");
            Test(-12345.67891, "# ??/100", "-12345 68/100");

            Test(1, "??/??", " 1/1 ");
            Test(-1.2, "??/??", "- 6/5 ");
            Test(12.3, "??/??", "123/10");
            Test(-12.34, "??/??", "-617/50");
            Test(123.45, "??/??", "2469/20");
            Test(-123.456, "??/??", "-7037/57");
            Test(1234.567, "??/??", "119753/97");
            Test(-1234.5678, "??/??", "-100000/81");
            Test(12345.6789, "??/??", "1000000/81");
            Test(-12345.67891, "??/??", "-1000000/81");

            Test(0.3, "# ?/?", " 2/7");
            Test(1.3, "# ?/?", "1 1/3");
            Test(2.3, "# ?/?", "2 2/7");

            // Not sure what/why ssf does here:
            // Test(0.123251512342345, "# ??/?????????", "  480894/3901729");
            // Test(0.123251512342345, "# ?? / ?????????", "  480894 / 3901729");
            // This implementation instead renders like this:
            Test(0.123251512342345, "# ??/?????????", " 480894/3901729  ");
            Test(0.123251512342345, "# ?? / ?????????", " 480894 / 3901729  ");

            Test(0, "0", "0");

        }

        void TestExponents(double value, string expected1, string expected2, string expected3, string expected4)
        {
            // value	#0.0E+0	##0.0E+0	###0.0E+0	####0.0E+0
            var result1 = Format(value, "#0.0E+0", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected1, result1);

            var result2 = Format(value, "##0.0E+0", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected2, result2);

            var result3 = Format(value, "###0.0E+0", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected3, result3);

            var result4 = Format(value, "####0.0E+0", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected4, result4);
        }

        void TestNumber(double value, string expected1, string expected2, string expected3, string expected4, string expected5)
        {
            // value	?.?	??.??	???.???	???.?0?	???.?#?
            var result1 = Format(value, "?.?", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected1, result1);

            var result2 = Format(value, "??.??", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected2, result2);

            var result3 = Format(value, "???.???", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected3, result3);

            var result4 = Format(value, "???.?0?", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected4, result4);

            var result5 = Format(value, "???.?#?", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected5, result5);
        }


        [TestMethod]
        public void NumberFormat_TestNumber()
        {
            TestNumber(0.0, " . ", "  .  ", "   .   ", "   . 0 ", "   .  ");
            TestNumber(0.1, " .1", "  .1 ", "   .1  ", "   .10 ", "   .1 ");
            TestNumber(0.12, " .1", "  .12", "   .12 ", "   .12 ", "   .12 ");
            TestNumber(0.123, " .1", "  .12", "   .123", "   .123", "   .123");

            TestNumber(1.0, "1. ", " 1.  ", "  1.   ", "  1. 0 ", "  1.  ");
            TestNumber(1.1, "1.1", " 1.1 ", "  1.1  ", "  1.10 ", "  1.1 ");
            TestNumber(1.12, "1.1", " 1.12", "  1.12 ", "  1.12 ", "  1.12 ");
            TestNumber(1.123, "1.1", " 1.12", "  1.123", "  1.123", "  1.123");
        }

        [TestMethod]
        public void NumberFormat_TestExponent()
        {
            TestExponents(-1.23457E-13, "-12.3E-14", "-123.5E-15", "-1234.6E-16", "-123.5E-15");
            TestExponents(-12345.6789, "-1.2E+4", "-12.3E+3", "-1.2E+4", "-12345.7E+0");

            TestExponents(1.23457E-13, "12.3E-14", "123.5E-15", "1234.6E-16", "123.5E-15");
            TestExponents(1.23457E-12, "1.2E-12", "1.2E-12", "1.2E-12", "1234.6E-15");
            TestExponents(1.23457E-11, "12.3E-12", "12.3E-12", "12.3E-12", "12345.7E-15");
            TestExponents(1.23457E-10, "1.2E-10", "123.5E-12", "123.5E-12", "1.2E-10");
            TestExponents(1.23457E-09, "12.3E-10", "1.2E-9", "1234.6E-12", "12.3E-10");
            TestExponents(1.23457E-08, "1.2E-8", "12.3E-9", "1.2E-8", "123.5E-10");
            TestExponents(0.000000123457, "12.3E-8", "123.5E-9", "12.3E-8", "1234.6E-10");
            TestExponents(0.00000123457, "1.2E-6", "1.2E-6", "123.5E-8", "12345.7E-10");
            TestExponents(0.0000123457, "12.3E-6", "12.3E-6", "1234.6E-8", "1.2E-5");
            TestExponents(0.000123457, "1.2E-4", "123.5E-6", "1.2E-4", "12.3E-5");
            TestExponents(0.001234568, "12.3E-4", "1.2E-3", "12.3E-4", "123.5E-5");
            TestExponents(0.012345679, "1.2E-2", "12.3E-3", "123.5E-4", "1234.6E-5");
            TestExponents(0.123456789, "12.3E-2", "123.5E-3", "1234.6E-4", "12345.7E-5");
            TestExponents(1.23456789, "1.2E+0", "1.2E+0", "1.2E+0", "1.2E+0");
            TestExponents(12.3456789, "12.3E+0", "12.3E+0", "12.3E+0", "12.3E+0");
            TestExponents(123.456789, "1.2E+2", "123.5E+0", "123.5E+0", "123.5E+0");
            TestExponents(1234.56789, "12.3E+2", "1.2E+3", "1234.6E+0", "1234.6E+0");
            TestExponents(12345.6789, "1.2E+4", "12.3E+3", "1.2E+4", "12345.7E+0");
            TestExponents(123456.789, "12.3E+4", "123.5E+3", "12.3E+4", "1.2E+5");
            TestExponents(1234567.89, "1.2E+6", "1.2E+6", "123.5E+4", "12.3E+5");
            TestExponents(12345678.9, "12.3E+6", "12.3E+6", "1234.6E+4", "123.5E+5");
            TestExponents(123456789D, "1.2E+8", "123.5E+6", "1.2E+8", "1234.6E+5");
            TestExponents(1234567890D, "12.3E+8", "1.2E+9", "12.3E+8", "12345.7E+5");
            TestExponents(12345678900D, "1.2E+10", "12.3E+9", "123.5E+8", "1.2E+10");
            TestExponents(123456789000D, "12.3E+10", "123.5E+9", "1234.6E+8", "12.3E+10");
            TestExponents(1234567890000D, "1.2E+12", "1.2E+12", "1.2E+12", "123.5E+10");
            TestExponents(12345678900000D, "12.3E+12", "12.3E+12", "12.3E+12", "1234.6E+10");
            TestExponents(123456789000000D, "1.2E+14", "123.5E+12", "123.5E+12", "12345.7E+10");
            TestExponents(1234567890000000D, "12.3E+14", "1.2E+15", "1234.6E+12", "1.2E+15");
            TestExponents(12345678900000000D, "1.2E+16", "12.3E+15", "1.2E+16", "12.3E+15");
            TestExponents(123456789000000000D, "12.3E+16", "123.5E+15", "12.3E+16", "123.5E+15");
            TestExponents(1234567890000000000D, "1.2E+18", "1.2E+18", "123.5E+16", "1234.6E+15");
            TestExponents(12345678900000000000D, "12.3E+18", "12.3E+18", "1234.6E+16", "12345.7E+15");
            TestExponents(123456789000000000000D, "1.2E+20", "123.5E+18", "1.2E+20", "1.2E+20");
            TestExponents(1234567890000000000000D, "12.3E+20", "1.2E+21", "12.3E+20", "12.3E+20");
            TestExponents(12345678900000000000000D, "1.2E+22", "12.3E+21", "123.5E+20", "123.5E+20");
            TestExponents(123456789000000000000000D, "12.3E+22", "123.5E+21", "1234.6E+20", "1234.6E+20");
            TestExponents(1234567890000000000000000D, "1.2E+24", "1.2E+24", "1.2E+24", "12345.7E+20");
            TestExponents(12345678900000000000000000D, "12.3E+24", "12.3E+24", "12.3E+24", "1.2E+25");
            TestExponents(123456789000000000000000000D, "1.2E+26", "123.5E+24", "123.5E+24", "12.3E+25");
            TestExponents(1234567890000000000000000000D, "12.3E+26", "1.2E+27", "1234.6E+24", "123.5E+25");
            TestExponents(12345678900000000000000000000D, "1.2E+28", "12.3E+27", "1.2E+28", "1234.6E+25");
            TestExponents(123456789000000000000000000000D, "12.3E+28", "123.5E+27", "12.3E+28", "12345.7E+25");
            TestExponents(1234567890000000000000000000000D, "1.2E+30", "1.2E+30", "123.5E+28", "1.2E+30");
            TestExponents(12345678900000000000000000000000D, "12.3E+30", "12.3E+30", "1234.6E+28", "12.3E+30");
        }

        void TestComma(double value, string expected1, string expected2, string expected3, string expected4, string expected5, string expected6, string expected7)
        {
            // value	#.0000,,,	#.0000,,	#.0000,	#,##0.0	###,##0	###,###	#,###.00
            var result1 = Format(value, "#.0000,,,", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected1, result1);

            var result2 = Format(value, "#.0000,,", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected2, result2);

            var result3 = Format(value, "#.0000,", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected3, result3);

            var result4 = Format(value, "#,##0.0", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected4, result4);

            var result5 = Format(value, "###,##0", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected5, result5);

            var result6 = Format(value, "###,###", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected6, result6);

            var result7 = Format(value, "#,###.00", CultureInfo.InvariantCulture);
            Assert.AreEqual(expected7, result7);
        }

        [TestMethod]
        public void NumberFormat_TestComma()
        {
            TestComma(0.99, ".0000", ".0000", ".0010", "1.0", "1", "1", ".99");
            TestComma(1.2345, ".0000", ".0000", ".0012", "1.2", "1", "1", "1.23");
            TestComma(12.345, ".0000", ".0000", ".0123", "12.3", "12", "12", "12.35");
            TestComma(123.456, ".0000", ".0001", ".1235", "123.5", "123", "123", "123.46");
            TestComma(1234, ".0000", ".0012", "1.2340", "1,234.0", "1,234", "1,234", "1,234.00");
            TestComma(12345, ".0000", ".0123", "12.3450", "12,345.0", "12,345", "12,345", "12,345.00");
            TestComma(123456, ".0001", ".1235", "123.4560", "123,456.0", "123,456", "123,456", "123,456.00");
            TestComma(1234567, ".0012", "1.2346", "1234.5670", "1,234,567.0", "1,234,567", "1,234,567", "1,234,567.00");
            TestComma(12345678, ".0123", "12.3457", "12345.6780", "12,345,678.0", "12,345,678", "12,345,678", "12,345,678.00");
            TestComma(123456789, ".1235", "123.4568", "123456.7890", "123,456,789.0", "123,456,789", "123,456,789", "123,456,789.00");
            TestComma(1234567890, "1.2346", "1234.5679", "1234567.8900", "1,234,567,890.0", "1,234,567,890", "1,234,567,890", "1,234,567,890.00");
            TestComma(12345678901, "12.3457", "12345.6789", "12345678.9010", "12,345,678,901.0", "12,345,678,901", "12,345,678,901", "12,345,678,901.00");
            TestComma(123456789012, "123.4568", "123456.7890", "123456789.0120", "123,456,789,012.0", "123,456,789,012", "123,456,789,012", "123,456,789,012.00");
            TestComma(4321, ".0000", ".0043", "4.3210", "4,321.0", "4,321", "4,321", "4,321.00");
            TestComma(4321234, ".0043", "4.3212", "4321.2340", "4,321,234.0", "4,321,234", "4,321,234", "4,321,234.00");

        }

        void TestValid(string format)
        {
            var to = new NumberFormatString(format);
            Assert.IsTrue(to.IsValid, "Invalid format: {0}", format);
        }

        [TestMethod]
        public void NumberFormat_TestValid()
        {
            TestValid("\" Excellent\"");
            TestValid("\" Fair\"");
            TestValid("\" Good\"");
            TestValid("\" Poor\"");
            TestValid("\" Very Good\"");
            TestValid("\"$\"#,##0");
            TestValid("\"$\"#,##0.00");
            TestValid("\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)");
            TestValid("\"$\"#,##0.00_);\\(\"$\"#,##0.00\\)");
            TestValid("\"$\"#,##0;[Red]\\-\"$\"#,##0");
            TestValid("\"$\"#,##0_);[Red]\\(\"$\"#,##0\\)");
            TestValid("\"$\"#,##0_);\\(\"$\"#,##0\\)");
            TestValid("\"Haha!\"\\ @\\ \"Yeah!\"");
            TestValid("\"TRUE\";\"TRUE\";\"FALSE\"");
            TestValid("\"True\";\"True\";\"False\";@");
            TestValid("\"Years: \"0");
            TestValid("\"Yes\";\"Yes\";\"No\";@");
            TestValid("\"kl \"hh:mm:ss;@");
            TestValid("\"£\"#,##0.00");
            TestValid("\"£\"#,##0;[Red]\\-\"£\"#,##0");
            TestValid("\"€\"#,##0.00");
            TestValid("\"€\"\\ #,##0.00_-");
            TestValid("\"上午/下午 \"hh\"時\"mm\"分\"ss\"秒 \"");
            TestValid("\"￥\"#,##0.00;\"￥\"\\-#,##0.00");
            TestValid("#");
            TestValid("# ?/?");
            TestValid("# ??/??");
            TestValid("#\" \"?/?");
            TestValid("#\" \"??/??");
            TestValid("#\"abded\"\\ ??/??");
            TestValid("###0.00;-###0.00");
            TestValid("###0;-###0");
            TestValid("##0.0E+0");
            TestValid("#,##0");
            TestValid("#,##0 ;(#,##0)");
            TestValid("#,##0 ;[Red](#,##0)");
            TestValid("#,##0\"р.\";[Red]\\-#,##0\"р.\"");
            TestValid("#,##0.0");
            TestValid("#,##0.00");
            TestValid("#,##0.00 \"�\"");
            TestValid("#,##0.00 €;-#,##0.00 €");
            TestValid("#,##0.00\"р.\";[Red]\\-#,##0.00\"р.\"");
            TestValid("#,##0.000");
            TestValid("#,##0.0000");
            TestValid("#,##0.00000");
            TestValid("#,##0.000000");
            TestValid("#,##0.0000000");
            TestValid("#,##0.00000000");
            TestValid("#,##0.000000000");
            TestValid("#,##0.00000000;[Red]#,##0.00000000");
            TestValid("#,##0.0000_ ");
            TestValid("#,##0.000_ ");
            TestValid("#,##0.000_);\\(#,##0.000\\)");
            TestValid("#,##0.00;(#,##0.00)");
            TestValid("#,##0.00;(#,##0.00);0.00");
            TestValid("#,##0.00;[Red](#,##0.00)");
            TestValid("#,##0.00;[Red]\\(#,##0.00\\)");
            TestValid("#,##0.00;\\(#,##0.00\\)");
            TestValid("#,##0.00[$₹-449]_);\\(#,##0.00[$₹-449]\\)");
            TestValid("#,##0.00\\ \"р.\"");
            TestValid("#,##0.00\\ \"р.\";[Red]\\-#,##0.00\\ \"р.\"");
            TestValid("#,##0.00\\ [$€-407]");
            TestValid("#,##0.00\\ [$€-40C]");
            TestValid("#,##0.00_);\\(#,##0.00\\)");
            TestValid("#,##0.00_р_.;[Red]\\-#,##0.00_р_.");
            TestValid("#,##0.00_р_.;\\-#,##0.00_р_.");
            TestValid("#,##0.0;[Red]#,##0.0");
            TestValid("#,##0.0_ ;\\-#,##0.0\\ ");
            TestValid("#,##0.0_);[Red]\\(#,##0.0\\)");
            TestValid("#,##0.0_);\\(#,##0.0\\)");
            TestValid("#,##0;\\-#,##0;0");
            TestValid("#,##0\\ \"р.\";[Red]\\-#,##0\\ \"р.\"");
            TestValid("#,##0\\ \"р.\";\\-#,##0\\ \"р.\"");
            TestValid("#,##0\\ ;[Red]\\(#,##0\\)");
            TestValid("#,##0\\ ;\\(#,##0\\)");
            TestValid("#,##0_ ");
            TestValid("#,##0_ ;[Red]\\-#,##0\\ ");
            TestValid("#,##0_);[Red]\\(#,##0\\)");
            TestValid("#,##0_р_.;[Red]\\-#,##0_р_.");
            TestValid("#,##0_р_.;\\-#,##0_р_.");
            TestValid("#.0000,,");
            TestValid("#0");
            TestValid("#0.00");
            TestValid("#0.0000");
            TestValid("#\\ ?/10");
            TestValid("#\\ ?/2");
            TestValid("#\\ ?/4");
            TestValid("#\\ ?/8");
            TestValid("#\\ ?/?");
            TestValid("#\\ ??/100");
            TestValid("#\\ ??/100;[Red]\\(#\\ ??/16\\)");
            TestValid("#\\ ??/16");
            TestValid("#\\ ??/??");
            TestValid("#\\ ??/?????????");
            TestValid("#\\ ???/???");
            TestValid("**\\ #,###,#00,000.00,**");
            TestValid("0");
            TestValid("0\"abde\".0\"??\"000E+00");
            TestValid("0%");
            TestValid("0.0");
            TestValid("0.0%");
            TestValid("0.00");
            TestValid("0.00\"°\"");
            TestValid("0.00%");
            TestValid("0.000");
            TestValid("0.000%");
            TestValid("0.0000");
            TestValid("0.000000");
            TestValid("0.00000000");
            TestValid("0.000000000");
            TestValid("0.000000000%");
            TestValid("0.00000000000");
            TestValid("0.000000000000000");
            TestValid("0.00000000E+00");
            TestValid("0.0000E+00");
            TestValid("0.00;[Red]0.00");
            TestValid("0.00E+00");
            TestValid("0.00_);[Red]\\(0.00\\)");
            TestValid("0.00_);\\(0.00\\)");
            TestValid("0.0_ ");
            TestValid("00.00.00.000");
            TestValid("00.000%");
            TestValid("0000");
            TestValid("00000");
            TestValid("00000000");
            TestValid("000000000");
            TestValid("00000\\-0000");
            TestValid("00000\\-00000");
            TestValid("000\\-00\\-0000");
            TestValid("0;[Red]0");
            TestValid("0\\-00000\\-00000\\-0");
            TestValid("0_);[Red]\\(0\\)");
            TestValid("0_);\\(0\\)");
            TestValid("@");
            TestValid("A/P");
            TestValid("AM/PM");
            TestValid("AM/PMh\"時\"mm\"分\"ss\"秒\";@");
            TestValid("D");
            TestValid("DD");
            TestValid("DD/MM/YY;@");
            TestValid("DD/MM/YYYY");
            TestValid("DD/MM/YYYY;@");
            TestValid("DDD");
            TestValid("DDDD");
            TestValid("DDDD\", \"MMMM\\ DD\", \"YYYY");
            TestValid("GENERAL");
            TestValid("General");
            TestValid("H");
            TestValid("H:MM:SS\\ AM/PM");
            TestValid("HH:MM");
            TestValid("HH:MM:SS\\ AM/PM");
            TestValid("HHM");
            TestValid("HHMM");
            TestValid("HH[MM]");
            TestValid("HH[M]");
            TestValid("M/D/YYYY");
            TestValid("M/D/YYYY\\ H:MM");
            TestValid("MM/DD/YY");
            TestValid("S");
            TestValid("SS");
            TestValid("YY");
            TestValid("YYM");
            TestValid("YYMM");
            TestValid("YYMMM");
            TestValid("YYMMMM");
            TestValid("YYMMMMM");
            TestValid("YYYY");
            TestValid("YYYY-MM-DD HH:MM:SS");
            TestValid("YYYY\\-MM\\-DD");
            TestValid("[$$-409]#,##0");
            TestValid("[$$-409]#,##0.00");
            TestValid("[$$-409]#,##0.00_);[Red]\\([$$-409]#,##0.00\\)");
            TestValid("[$$-C09]#,##0.00");
            TestValid("[$-100042A]h:mm:ss\\ AM/PM;@");
            TestValid("[$-1010409]0.000%");
            TestValid("[$-1010409]General");
            TestValid("[$-1010409]d/m/yyyy\\ h:mm\\ AM/PM;@");
            TestValid("[$-1010409]dddd, mmmm dd, yyyy");
            TestValid("[$-1010409]m/d/yyyy");
            TestValid("[$-1409]h:mm:ss\\ AM/PM;@");
            TestValid("[$-2000000]h:mm:ss;@");
            TestValid("[$-2010401]d/mm/yyyy\\ h:mm\\ AM/PM;@");
            TestValid("[$-4000439]h:mm:ss\\ AM/PM;@");
            TestValid("[$-4010439]d/m/yyyy\\ h:mm\\ AM/PM;@");
            TestValid("[$-409]AM/PM\\ hh:mm:ss;@");
            TestValid("[$-409]d/m/yyyy\\ hh:mm;@");
            TestValid("[$-409]d\\-mmm;@");
            TestValid("[$-409]d\\-mmm\\-yy;@");
            TestValid("[$-409]d\\-mmm\\-yyyy;@");
            TestValid("[$-409]dd/mm/yyyy\\ hh:mm;@");
            TestValid("[$-409]dd\\-mmm\\-yy;@");
            TestValid("[$-409]h:mm:ss\\ AM/PM;@");
            TestValid("[$-409]h:mm\\ AM/PM;@");
            TestValid("[$-409]m/d/yy\\ h:mm\\ AM/PM;@");
            TestValid("[$-409]mmm\\-yy;@");
            TestValid("[$-409]mmmm\\ d\\,\\ yyyy;@");
            TestValid("[$-409]mmmm\\-yy;@");
            TestValid("[$-409]mmmmm;@");
            TestValid("[$-409]mmmmm\\-yy;@");
            TestValid("[$-40E]h\\ \"óra\"\\ m\\ \"perckor\"\\ AM/PM;@");
            TestValid("[$-412]AM/PM\\ h\"시\"\\ mm\"분\"\\ ss\"초\";@");
            TestValid("[$-41C]h:mm:ss\\.AM/PM;@");
            TestValid("[$-449]hh:mm:ss\\ AM/PM;@");
            TestValid("[$-44E]hh:mm:ss\\ AM/PM;@");
            TestValid("[$-44F]hh:mm:ss\\ AM/PM;@");
            TestValid("[$-D000409]h:mm\\ AM/PM;@");
            TestValid("[$-D010000]d/mm/yyyy\\ h:mm\\ \"น.\";@");
            TestValid("[$-F400]h:mm:ss\\ AM/PM");
            TestValid("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy");
            TestValid("[$AUD]\\ #,##0.00");
            TestValid("[$RD$-1C0A]#,##0.00;[Red]\\-[$RD$-1C0A]#,##0.00");
            TestValid("[$SFr.-810]\\ #,##0.00_);[Red]\\([$SFr.-810]\\ #,##0.00\\)");
            TestValid("[$£-809]#,##0.00;[Red][$£-809]#,##0.00");
            TestValid("[$¥-411]#,##0.00");
            TestValid("[$¥-804]#,##0.00");
            TestValid("[<0]\"\";0%");
            TestValid("[<=9999999]###\\-####;\\(###\\)\\ ###\\-####");
            TestValid("[=0]?;#,##0.00");
            TestValid("[=0]?;0%");
            TestValid("[=0]?;[<4.16666666666667][hh]:mm:ss;[hh]:mm");
            TestValid("[>999999]#,,\"M\";[>999]#,\"K\";#");
            TestValid("[>999999]#.000,,\"M\";[>999]#.000,\"K\";#.000");
            TestValid("[>=100000]0.000\\ \\\";[Red]0.000\\ \\<\\ \\>\\ \\\"\\ \\&\\ \\'\\ ");
            TestValid("[>=100000]0.000\\ \\<;[Red]0.000\\ \\>");
            TestValid("[BLACK]@");
            TestValid("[BLUE]GENERAL");
            TestValid("[Black]@");
            TestValid("[Blue]General");
            TestValid("[CYAN]@");
            TestValid("[Cyan]@");
            TestValid("[DBNum1][$-804]AM/PMh\"时\"mm\"分\";@");
            TestValid("[DBNum1][$-804]General");
            TestValid("[DBNum1][$-804]h\"时\"mm\"分\";@");
            TestValid("[ENG][$-1004]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@");
            TestValid("[ENG][$-101040D]d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-101042A]d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-140C]dddd\\ \"YeahWoo!\"\\ ddd\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-2C0A]dddd\\ d\" de \"mmmm\" de \"yyyy;@");
            TestValid("[ENG][$-402]dd\\ mmmm\\ yyyy\\ \"г.\";@");
            TestValid("[ENG][$-403]dddd\\,\\ d\" / \"mmmm\" / \"yyyy;@");
            TestValid("[ENG][$-405]d\\.\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-408]d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-409]d\\-mmm;@");
            TestValid("[ENG][$-409]d\\-mmm\\-yy;@");
            TestValid("[ENG][$-409]d\\-mmm\\-yyyy;@");
            TestValid("[ENG][$-409]dd\\-mmm\\-yy;@");
            TestValid("[ENG][$-409]mmm\\-yy;@");
            TestValid("[ENG][$-409]mmmm\\ d\\,\\ yyyy;@");
            TestValid("[ENG][$-409]mmmm\\-yy;@");
            TestValid("[ENG][$-40B]d\\.\\ mmmm\\t\\a\\ yyyy;@");
            TestValid("[ENG][$-40C]d/mmm/yyyy;@");
            TestValid("[ENG][$-40E]yyyy/\\ mmmm\\ d\\.;@");
            TestValid("[ENG][$-40F]dd\\.\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-410]d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-415]d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-416]d\\ \\ mmmm\\,\\ yyyy;@");
            TestValid("[ENG][$-418]d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-41A]d\\.\\ mmmm\\ yyyy\\.;@");
            TestValid("[ENG][$-41B]d\\.\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-41D]\"den \"\\ d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-420]dddd\\,\\ dd\\ mmmm\\,\\ yyyy;@");
            TestValid("[ENG][$-421]dd\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-424]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-425]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-426]dddd\\,\\ yyyy\". gada \"d\\.\\ mmmm;@");
            TestValid("[ENG][$-427]yyyy\\ \"m.\"\\ mmmm\\ d\\ \"d.\";@");
            TestValid("[ENG][$-42B]dddd\\,\\ d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-42C]d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-42D]yyyy\"(e)ko\"\\ mmmm\"ren\"\\ d\"a\";@");
            TestValid("[ENG][$-42F]dddd\\,\\ dd\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-437]yyyy\\ \\წ\\ლ\\ი\\ს\\ dd\\ mm\\,\\ dddd;@");
            TestValid("[ENG][$-438]d\\.\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-43F]d\\ mmmm\\ yyyy\\ \"ж.\";@");
            TestValid("[ENG][$-444]d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-449]dd\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-44E]d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-44F]dd\\ mmmm\\ yyyy\\ dddd;@");
            TestValid("[ENG][$-457]dd\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-813]dddd\\ d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-81A]dddd\\,\\ d\\.\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-82C]d\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-843]yyyy\\ \"й\"\"и\"\"л\"\\ d/mmmm;@");
            TestValid("[ENG][$-C07]dddd\\,\\ dd\\.\\ mmmm\\ yyyy;@");
            TestValid("[ENG][$-FC19]yyyy\\,\\ dd\\ mmmm;@");
            TestValid("[ENG][$-FC22]d\\ mmmm\\ yyyy\" р.\";@");
            TestValid("[ENG][$-FC23]d\\ mmmm\\ yyyy;@");
            TestValid("[GREEN]#,###");
            TestValid("[Green]#,###");
            TestValid("[HH]");
            TestValid("[HIJ][$-2060401]d/mm/yyyy\\ h:mm\\ AM/PM;@");
            TestValid("[HIJ][$-2060401]d\\ mmmm\\ yyyy;@");
            TestValid("[H]");
            TestValid("[JPN][$-411]gggyy\"年\"m\"月\"d\"日\"\\ dddd;@");
            TestValid("[MAGENTA]0.00");
            TestValid("[Magenta]0.00");
            TestValid("[RED]#.##");
            TestValid("[Red]#.##");
            TestValid("[Red][<-25]General;[Blue][>25]General;[Green]General;[Yellow]General\\ ");
            TestValid("[Red][<=-25]General;[Blue][>=25]General;[Green]General;[Yellow]General");
            TestValid("[Red][<>50]General;[Blue]000");
            TestValid("[Red][=50]General;[Blue]000");
            TestValid("[SS]");
            TestValid("[S]");
            TestValid("[TWN][DBNum1][$-404]y\"年\"m\"月\"d\"日\";@");
            TestValid("[WHITE]0.0");
            TestValid("[White]0.0");
            TestValid("[YELLOW]@");
            TestValid("[Yellow]@");
            TestValid("[h]");
            TestValid("[h]:mm:ss");
            TestValid("[h]:mm:ss;@");
            TestValid("[h]\\.mm\" Uhr \";@");
            TestValid("[hh]");
            TestValid("[s]");
            TestValid("[ss]");
            TestValid("\\#\\r\\e\\c");
            TestValid("\\$#,##0_);[Red]\"($\"#,##0\\)");
            TestValid("\\$0.00");
            TestValid("\\C\\O\\B\\ \\o\\n\\ @");
            TestValid("\\C\\R\\O\\N\\T\\A\\B\\ \\o\\n\\ @");
            TestValid("\\R\\e\\s\\u\\l\\t\\ \\o\\n\\ @");
            TestValid("\\S\\Q\\L\\ \\:\\ @");
            TestValid("\\S\\Q\\L\\ \\R\\e\\q\\u\\e\\s\\t\\ \\f\\o\\r\\ @");
            TestValid("\\c\\c\\c?????0\"aaaa\"0\"bbbb\"000000.00%");
            TestValid("\\u\\n\\t\\i\\l\\ h:mm;@");
            TestValid("_ \"￥\"* #,##0.00_ \"Positive\";_ \"￥\"* \\-#,##0.00_ ;_ \"￥\"* \"-\"??_ \"Negtive\";_ @_ \\ \"Zero\"");
            TestValid("_ * #,##0.00_)[$﷼-429]_ ;_ * \\(#,##0.00\\)[$﷼-429]_ ;_ * \"-\"??_)[$﷼-429]_ ;_ @_ ");
            TestValid("_ * #,##0_ ;_ * \\-#,##0_ ;[Red]_ * \"-\"_ ;_ @_ ");
            TestValid("_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)");
            TestValid("_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"??_);_(@_)");
            TestValid("_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)");
            TestValid("_(* #,##0.0000_);_(* \\(#,##0.0000\\);_(* \"-\"??_);_(@_)");
            TestValid("_(* #,##0.000_);_(* \\(#,##0.000\\);_(* \"-\"??_);_(@_)");
            TestValid("_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)");
            TestValid("_(* #,##0.0_);_(* \\(#,##0.0\\);_(* \"-\"??_);_(@_)");
            TestValid("_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"??_);_(@_)");
            TestValid("_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)");
            TestValid("_([$ANG]\\ * #,##0.0_);_([$ANG]\\ * \\(#,##0.0\\);_([$ANG]\\ * \"-\"?_);_(@_)");
            TestValid("_-\"€\"\\ * #,##0.00_-;_-\"€\"\\ * #,##0.00\\-;_-\"€\"\\ * \"-\"??_-;_-@_-");
            TestValid("_-* #,##0.00\" TL\"_-;\\-* #,##0.00\" TL\"_-;_-* \\-??\" TL\"_-;_-@_-");
            TestValid("_-* #,##0.00\" €\"_-;\\-* #,##0.00\" €\"_-;_-* \\-??\" €\"_-;_-@_-");
            TestValid("_-* #,##0.00\\ \"р.\"_-;\\-* #,##0.00\\ \"р.\"_-;_-* \"-\"??\\ \"р.\"_-;_-@_-");
            TestValid("_-* #,##0.00\\ \"€\"_-;\\-* #,##0.00\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_-");
            TestValid("_-* #,##0.00\\ [$€-407]_-;\\-* #,##0.00\\ [$€-407]_-;_-* \\-??\\ [$€-407]_-;_-@_-");
            TestValid("_-* #,##0.0\\ _F_-;\\-* #,##0.0\\ _F_-;_-* \"-\"??\\ _F_-;_-@_-");
            TestValid("_-* #,##0\\ \"€\"_-;\\-* #,##0\\ \"€\"_-;_-* \"-\"\\ \"€\"_-;_-@_-");
            TestValid("_-* #,##0_-;\\-* #,##0_-;_-* \"-\"??_-;_-@_-");
            TestValid("_-\\$* #,##0.0_ ;_-\\$* \\-#,##0.0\\ ;_-\\$* \"-\"?_ ;_-@_ ");
            TestValid("d");
            TestValid("d-mmm");
            TestValid("d-mmm-yy");
            TestValid("d/m");
            TestValid("d/m/yy;@");
            TestValid("d/m/yyyy;@");
            TestValid("d/mm/yy;@");
            TestValid("d/mm/yyyy;@");
            TestValid("d\\-mmm");
            TestValid("d\\-mmm\\-yyyy");
            TestValid("dd");
            TestValid("dd\"-\"mmm\"-\"yyyy");
            TestValid("dd/m/yyyy");
            TestValid("dd/mm/yy");
            TestValid("dd/mm/yy;@");
            TestValid("dd/mm/yy\\ hh:mm");
            TestValid("dd/mm/yyyy");
            TestValid("dd/mm/yyyy\\ hh:mm:ss");
            TestValid("dd/mmm");
            TestValid("dd\\-mm\\-yy");
            TestValid("dd\\-mmm\\-yy");
            TestValid("dd\\-mmm\\-yyyy\\ hh:mm:ss.000");
            TestValid("dd\\/mm\\/yy");
            TestValid("dd\\/mm\\/yyyy");
            TestValid("ddd");
            TestValid("dddd");
            TestValid("dddd, mmmm dd, yyyy");
            TestValid("h");
            TestValid("h\"时\"mm\"分\"ss\"秒\";@");
            TestValid("h\"時\"mm\"分\"ss\"秒\";@");
            TestValid("h:mm");
            TestValid("h:mm AM/PM");
            TestValid("h:mm:ss");
            TestValid("h:mm:ss AM/PM");
            TestValid("h:mm:ss;@");
            TestValid("h:mm;@");
            TestValid("h\\.mm\" Uhr \";@");
            TestValid("h\\.mm\" h\";@");
            TestValid("h\\.mm\" u.\";@");
            TestValid("hh\":\"mm AM/PM");
            TestValid("hh:mm:ss");
            TestValid("hh:mm:ss\\ AM/PM");
            TestValid("hh\\.mm\" h\";@");
            TestValid("hhm");
            TestValid("hhmm");
            TestValid("m\"月\"d\"日\"");
            TestValid("m/d/yy");
            TestValid("m/d/yy h:mm");
            TestValid("m/d/yy;@");
            TestValid("m/d/yy\\ h:mm");
            TestValid("m/d/yy\\ h:mm;@");
            TestValid("m/d/yyyy");
            TestValid("m/d/yyyy;@");
            TestValid("m/d/yyyy\\ h:mm:ss;@");
            TestValid("m/d;@");
            TestValid("m\\/d\\/yyyy");
            TestValid("mm/dd");
            TestValid("mm/dd/yy");
            TestValid("mm/dd/yy;@");
            TestValid("mm/dd/yyyy");
            TestValid("mm:ss");
            TestValid("mm:ss.0;@");
            TestValid("mmm d, yyyy");
            TestValid("mmm\" \"d\", \"yyyy");
            TestValid("mmm-yy");
            TestValid("mmm-yy;@");
            TestValid("mmm/yy");
            TestValid("mmm\\-yy");
            TestValid("mmm\\-yy;@");
            TestValid("mmm\\-yyyy");
            TestValid("mmmm\\ d\\,\\ yyyy");
            TestValid("mmmm\\ yyyy");
            TestValid("mmss.0");
            TestValid("s");
            TestValid("ss");
            TestValid("yy");
            TestValid("yy/mm/dd");
            TestValid("yy\\.mm\\.dd");
            TestValid("yym");
            TestValid("yymm");
            TestValid("yymmm");
            TestValid("yymmmm");
            TestValid("yymmmmm");
            TestValid("yyyy");
            TestValid("yyyy\"년\"\\ m\"월\"\\ d\"일\";@");
            TestValid("yyyy-m-d h:mm AM/PM");
            TestValid("yyyy-mm-dd");
            TestValid("yyyy/mm/dd");
            TestValid("yyyy\\-m\\-d\\ hh:mm:ss");
            TestValid("yyyy\\-mm\\-dd");
            TestValid("yyyy\\-mm\\-dd;@");
            TestValid("yyyy\\-mm\\-dd\\ h:mm");
            TestValid("yyyy\\-mm\\-dd\\Thh:mm");
            TestValid("yyyy\\-mm\\-dd\\Thhmmss.000");
        }
    }
}
