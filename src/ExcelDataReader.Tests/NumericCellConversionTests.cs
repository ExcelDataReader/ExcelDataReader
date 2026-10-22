using ExcelDataReader.Core;

namespace ExcelDataReader.Tests;

public class NumericCellConversionTests
{
    [TestCase(false)]
    [TestCase(true)]
    public void NumericValuesPreserveTheirOriginalType(bool date1904)
    {
        var workbook = new CommonWorkbook();
        Assert.That(workbook.ConvertNumericValue(42, 0, date1904), Is.TypeOf<int>().And.EqualTo(42));
        Assert.That(workbook.ConvertNumericValue(42.5, 0, date1904), Is.TypeOf<double>().And.EqualTo(42.5));
        Assert.That(workbook.ConvertNumericValue(3000000, 14, date1904), Is.TypeOf<int>().And.EqualTo(3000000));
        Assert.That(workbook.ConvertNumericValue(3000000.5, 14, date1904), Is.TypeOf<double>().And.EqualTo(3000000.5));
        Assert.That(double.IsNaN((double)workbook.ConvertNumericValue(double.NaN, 14, date1904)), Is.True);
        Assert.That(workbook.ConvertNumericValue(double.PositiveInfinity, 14, date1904), Is.EqualTo(double.PositiveInfinity));
    }

    [TestCase(0, false, 1899, 12, 31)]
    [TestCase(1, false, 1900, 1, 1)]
    [TestCase(59, false, 1900, 2, 28)]
    [TestCase(60, false, 1900, 2, 28)]
    [TestCase(61, false, 1900, 3, 1)]
    [TestCase(0, true, 1904, 1, 1)]
    [TestCase(1, true, 1904, 1, 2)]
    public void IntegerAndDoubleDatesUseTheSameEpoch(int serial, bool date1904, int year, int month, int day)
    {
        var workbook = new CommonWorkbook();
        var expected = new DateTime(year, month, day);
        Assert.That(workbook.ConvertNumericValue(serial, 14, date1904), Is.EqualTo(expected));
        Assert.That(workbook.ConvertNumericValue((double)serial, 14, date1904), Is.EqualTo(expected));
        Assert.That(workbook.ConvertNumericValue(serial + 0.5, 14, date1904), Is.EqualTo(expected.AddHours(12)));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void CustomFormatsOverrideBuiltinsAndDateClassificationPrecedesDuration(bool date1904)
    {
        var workbook = new CommonWorkbook();
        workbook.AddNumberFormat(14, "[h]:mm");
        workbook.AddNumberFormat(164, "yyyy;[h]:mm");
        Assert.That(workbook.ConvertNumericValue(2, 14, date1904), Is.EqualTo(TimeSpan.FromDays(2)));
        Assert.That(workbook.ConvertNumericValue(-0.5, 14, date1904), Is.EqualTo(TimeSpan.FromHours(-12)));
        Assert.Throws<OverflowException>(() => workbook.ConvertNumericValue(double.MaxValue, 14, date1904));
        Assert.Throws<OverflowException>(() => workbook.ConvertNumericValue(int.MaxValue, 14, date1904));
        Assert.That(
            workbook.ConvertNumericValue(2, 164, date1904),
            Is.EqualTo(date1904 ? new DateTime(1904, 1, 3) : new DateTime(1900, 1, 2)));

        var numericWorkbook = new CommonWorkbook();
        numericWorkbook.AddNumberFormat(14, "0.00");
        Assert.That(numericWorkbook.ConvertNumericValue(2, 14, date1904), Is.TypeOf<int>().And.EqualTo(2));
        Assert.That(numericWorkbook.ConvertNumericValue(2.5, 14, date1904), Is.TypeOf<double>().And.EqualTo(2.5));
    }

    [TestCase("xlsb", false)]
    [TestCase("xlsb", true)]
    [TestCase("xlsx", false)]
    [TestCase("xlsx", true)]
    public void DateAndDurationFixturesMatchXlsAfterEarlyAndCompleteReset(string extension, bool singlePass)
    {
        using var expectedReader = ExcelReaderFactory.CreateReader(
            Configuration.GetTestWorkbook("Issue283_TimeSpan.xls"),
            new ExcelReaderConfiguration { SinglePassMode = singlePass });
        using var actualReader = ExcelReaderFactory.CreateReader(
            Configuration.GetTestWorkbook("Issue283_TimeSpan." + extension),
            new ExcelReaderConfiguration { SinglePassMode = singlePass });

        Assert.That(expectedReader.Read(), Is.True);
        Assert.That(actualReader.Read(), Is.True);
        expectedReader.Reset();
        actualReader.Reset();

        for (int pass = 0; pass < 2; pass++)
        {
            int rows = 0;
            while (expectedReader.Read())
            {
                Assert.That(actualReader.Read(), Is.True);
                for (int column = 0; column < 3; column++)
                {
                    Assert.That(actualReader.GetValue(column), Is.EqualTo(expectedReader.GetValue(column)));
                    Assert.That(actualReader.GetFieldType(column), Is.EqualTo(expectedReader.GetFieldType(column)));
                    Assert.That(actualReader.GetCellError(column), Is.EqualTo(expectedReader.GetCellError(column)));
                }

                rows++;
            }

            Assert.That(rows, Is.GreaterThanOrEqualTo(9));

            // Open XML fixtures retain trailing styled rows in single-pass mode.
            while (actualReader.Read())
            {
                for (int column = 0; column < 3; column++)
                    Assert.That(actualReader.IsDBNull(column), Is.True);
            }

            Assert.That(actualReader.NextResult(), Is.EqualTo(expectedReader.NextResult()));
            expectedReader.Reset();
            actualReader.Reset();
        }
    }
}
