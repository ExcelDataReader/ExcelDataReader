namespace ExcelDataReader.Tests;

[TestFixture]
public class ExcelSpreadsheetXmlReaderTest
{
    [Test]
    public void ReadSpreadsheetXml_WorksheetsRowsAndTypes()
    {
        using var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("SpreadsheetXml2003.xml"));

        Assert.That(reader.ResultsCount, Is.EqualTo(2));
        Assert.That(reader.Name, Is.EqualTo("Sheet1"));
        Assert.That(reader.FieldCount, Is.EqualTo(4));

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("Name"));
        Assert.That(reader.GetString(1), Is.EqualTo("Value"));
        Assert.That(reader.GetString(2), Is.EqualTo("When"));
        Assert.That(reader.GetString(3), Is.EqualTo("Empty"));

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("A"));
        Assert.That(reader.GetDouble(1), Is.EqualTo(42D));
        Assert.That(reader.GetDateTime(2), Is.EqualTo(new DateTime(2024, 1, 2)));
        Assert.That(reader.IsDBNull(3), Is.True);
        Assert.That(reader.GetNumberFormatString(2), Is.EqualTo("yyyy-mm-dd"));

        // Missing row from ss:Index should still be represented.
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.IsDBNull(0), Is.True);
        Assert.That(reader.IsDBNull(1), Is.True);
        Assert.That(reader.IsDBNull(2), Is.True);
        Assert.That(reader.IsDBNull(3), Is.True);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetBoolean(1), Is.True);

        Assert.That(reader.NextResult(), Is.True);
        Assert.That(reader.Name, Is.EqualTo("Sheet2"));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("X"));
    }

    [Test]
    public void ReadSpreadsheetXml_LeadingWhitespaceBeforeDeclaration_IsTolerated()
    {
        using var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("SpreadsheetXml2003_LeadingWhitespace.xml"));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("ok"));
    }

    [Test]
    public void ReadSpreadsheetXml_HiddenRow_UsesZeroRowHeight()
    {
        using var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("SpreadsheetXml2003_HiddenRefreshRow.xml"));

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("Refresh"));
        Assert.That(reader.GetString(1), Is.EqualTo("<root><v>1</v></root>"));
        Assert.That(reader.RowHeight, Is.EqualTo(0D));

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("Visible"));
        Assert.That(reader.GetDouble(1), Is.EqualTo(1D));
        Assert.That(reader.RowHeight, Is.EqualTo(15D));
    }

    [Test]
    public void ReadSpreadsheetXml_SinglePassMode_RowCountThrowsAndFieldCountGrows()
    {
        using var reader = ExcelReaderFactory.CreateReader(
            Configuration.GetTestWorkbook("SpreadsheetXml2003_SinglePassFieldCount.xml"),
            new ExcelReaderConfiguration { SinglePassMode = true });

        Assert.That(reader.FieldCount, Is.Zero);
        Assert.Throws<InvalidOperationException>(() => _ = reader.RowCount);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("A"));
        Assert.That(reader.FieldCount, Is.EqualTo(1));

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(2), Is.EqualTo("C"));
        Assert.That(reader.FieldCount, Is.EqualTo(3));
    }

    [Test]
    public void ReadSpreadsheetXml_ScanMode_UsesActualRowAndColumnCounts()
    {
        using var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("SpreadsheetXml2003_WrongExpandedCounts.xml"));

        Assert.That(reader.FieldCount, Is.EqualTo(2));
        Assert.That(reader.RowCount, Is.EqualTo(2));
    }
}
