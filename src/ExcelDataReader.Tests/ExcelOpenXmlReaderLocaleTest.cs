using System.Data;
using System.Globalization;

namespace ExcelDataReader.Tests;

public class ExcelOpenXmlReaderLocaleTest
{
    [Test]
    public void TimeIsReadableForPolishLocaleIssueXxx()
    {
        Thread.CurrentThread.CurrentCulture = new CultureInfo("pl-PL", false);

        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Issue_xxx_LocaleTime.xlsx"));
        var dataSet = reader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[1][1], Is.EqualTo(new System.DateTime(1899, 12, 31, 1, 34, 0)));
        Assert.That(dataSet.Tables[0].Rows[2][1], Is.EqualTo(new System.DateTime(1899, 12, 31, 1, 34, 0)));
        Assert.That(dataSet.Tables[0].Rows[3][1], Is.EqualTo(new System.DateTime(1899, 12, 31, 18, 47, 0)));

        reader.Close();
    }

    [Test]
    public void TestDecimalLocale()
    {
        // change culture to german. this will expect commas instead of decimal points
        Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE", false);

        IExcelDataReader excelReader =
            ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_Decimal_Locale.xlsx"));

        var dataSet = excelReader.AsDataSet();

        excelReader.Close();

        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo(0.01));
        Assert.That(dataSet.Tables[0].Rows[1][0], Is.EqualTo(0.0001));
        Assert.That(dataSet.Tables[0].Rows[2][0], Is.EqualTo(0.123456789));
        Assert.That(dataSet.Tables[0].Rows[3][0], Is.EqualTo(0.00000000001));
    }

    [Test]
    //// [SetCulture("sv-SE")]
    public void CellFormat49()
    {
        Thread.CurrentThread.CurrentCulture = new CultureInfo("sv-SE", false);

        using IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Format49_@.xlsx"));
        DataSet result = excelReader.AsDataSet();

        // ExcelDataReader used to convert numbers formatted with NumFmtId=49/@ to culture-specific strings.
        // This behaviour changed in v3 to return the original value:
        // Assert.That(result.Tables[0].Rows[0].ItemArray, Is.EqualTo(new[] { "2010-05-05", "1.1", "2,2", "123", "2,2" }));
        Assert.That(result.Tables[0].Rows[0].ItemArray, Is.EqualTo(new object[] { "2010-05-05", "1.1", 2.2000000000000002D, 123.0D, "2,2" }));
    }
}
