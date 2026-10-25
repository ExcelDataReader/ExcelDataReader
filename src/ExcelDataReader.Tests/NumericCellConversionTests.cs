using System.Globalization;
using System.IO.Compression;
using ExcelDataReader.Core;
using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Tests;

public class NumericCellConversionTests
{
    [TestCase("0")]
    [TestCase("0000000000")]
    [TestCase("999999999")]
    [TestCase("2147483647")]
    [TestCase("2147483648")]
    [TestCase("1.25")]
    [TestCase("1e3")]
    [TestCase("-0")]
    [TestCase(" +42 ")]
    [TestCase("1,234")]
    [TestCase("42.0")]
    [TestCase("not-number")]
    public void XmlNumericParsingPreservesGeneralParserResults(string value)
    {
        using var reader = OpenXmlCell(value, string.Empty);
        Assert.That(reader.Read(), Is.True);
        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double expected))
        {
            Assert.That(reader.GetValue(0), Is.TypeOf<double>().And.EqualTo(expected));
            Assert.That(BitConverter.DoubleToInt64Bits(reader.GetDouble(0)), Is.EqualTo(BitConverter.DoubleToInt64Bits(expected)));
        }
        else
        {
            Assert.That(reader.GetString(0), Is.EqualTo(value));
        }
    }

    [TestCase("0")]
    [TestCase("0000000000")]
    [TestCase(" +0 ")]
    [TestCase("0.0")]
    [TestCase("0e0")]
    [TestCase("0,000")]
    [TestCase("-1")]
    [TestCase("2147483648")]
    public void XmlSharedStringIndicesPreserveGeneralParserResults(string value)
    {
        using var reader = OpenXmlCell(value, "s");
        Assert.That(reader.Read(), Is.True);
        if (int.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out int index))
            Assert.That(reader.GetValue(0), Is.EqualTo(index == 0 ? "shared" : DBNull.Value));
        else
            Assert.That(reader.GetString(0), Is.EqualTo(value));
    }

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

    private static IExcelDataReader OpenXmlCell(string value, string type)
    {
        var buffer = new MemoryStream();
        byte[] template = AllocationTestWorkbook.CreateXlsx(["<t>shared</t>"], true);
        buffer.Write(template, 0, template.Length);
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            archive.GetEntry("xl/worksheets/sheet1.xml")!.Delete();
            using var writer = new StreamWriter(archive.CreateEntry("xl/worksheets/sheet1.xml").Open());
            writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"" +
                type + "\"><v>" + value + "</v></c></row></sheetData></worksheet>");
        }

        buffer.Position = 0;
        return ExcelReaderFactory.CreateReader(buffer);
    }
}