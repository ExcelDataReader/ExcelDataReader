using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Tests;

public class CellStyleTests
{
    [TestCase("xls")]
    [TestCase("xlsx")]
    [TestCase("xlsb")]
    public void ReturnedStylesRemainStableAcrossReadingResetAndDisposal(string extension)
    {
        using var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("NumDoubleDateBoolString." + extension));
        Assert.That(reader.Read(), Is.True);
        var style = reader.GetCellStyle(0);
        var expected = (style.FontIndex, style.NumberFormatIndex, style.IndentLevel,
            style.HorizontalAlignment, style.VerticalAlignment, style.Hidden, style.Locked);
        Assert.That(reader.GetCellStyle(0), Is.SameAs(style));
        while (reader.Read())
        {
            for (int column = 0; column < reader.FieldCount; column++)
                _ = reader.GetCellStyle(column);
        }

        reader.Reset();
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetCellStyle(0), Is.SameAs(style));
        reader.Dispose();
        var actual = (style.FontIndex, style.NumberFormatIndex, style.IndentLevel,
            style.HorizontalAlignment, style.VerticalAlignment, style.Hidden, style.Locked);
        Assert.That(actual, Is.EqualTo(expected));
    }

    [Test]
    public void MissingCellsShareTheDefaultStyle()
    {
        const string xml = "<Data ss:Type=\"String\">value</Data>";
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(
            AllocationTestWorkbook.CreateSpreadsheetXml([xml, string.Empty, string.Empty])));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.Read(), Is.True);
        var style = reader.GetCellStyle(0);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetCellStyle(0), Is.SameAs(style));
        Assert.That(style.NumberFormatIndex, Is.Zero);
    }
}