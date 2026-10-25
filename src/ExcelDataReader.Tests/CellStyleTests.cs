using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Tests;

public class CellStyleTests
{
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