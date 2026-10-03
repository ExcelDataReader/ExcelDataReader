using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Tests;

[TestFixture]
public class ExcelSpreadsheetXmlReaderTest : ExcelSpreadsheetContractTestBase
{
    protected override DateTime Issue82_TodayDate => new(2013, 4, 19);

    protected override bool SupportsCodeName => false;

    [TestCase(false)]
    [TestCase(true)]
    public void DataTextFragments_PreserveContentAndReaderPosition(bool singlePass)
    {
        string[] data =
        [
            "<Data ss:Type=\"String\">single</Data>",
            "<Data ss:Type=\"String\"> \t\npreserved\r </Data>",
            "<Data ss:Type=\"String\">one<![CDATA[ two]]><B> three</B> four</Data>",
            "<Data ss:Type=\"String\"/>",
            "<Data ss:Type=\"Number\">12<![CDATA[.5]]></Data>",
            "<Data ss:Type=\"Boolean\">1</Data>",
            "<Data ss:Type=\"DateTime\">2026-01-02T03:04:05</Data>",
            "<Data ss:Type=\"Error\">#DIV/0!</Data>",
            "<Data ss:Type=\"String\">" + new string('\u03BB', 10000) + "</Data>",
            "<Data ss:Type=\"String\">last</Data>",
        ];
        object[] expected = ["single", " \t\npreserved\n ", "one two three four", string.Empty,
            12.5D, true, new DateTime(2026, 1, 2, 3, 4, 5), DBNull.Value, new string('\u03BB', 10000), "last"];
        using var reader = ExcelReaderFactory.CreateReader(
            new MemoryStream(AllocationTestWorkbook.CreateSpreadsheetXml(data)),
            new ExcelReaderConfiguration { SinglePassMode = singlePass });
        foreach (object value in expected)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetValue(0), Is.EqualTo(value));
            Assert.That(reader.GetFieldType(0), Is.EqualTo(value.GetType()));
            Assert.That(reader.GetCellError(0), Is.EqualTo(value is DBNull ? CellError.DIV0 : (CellError?)null));
        }

        Assert.That(reader.Read(), Is.False);
    }

    [Test]
    public void ReadSpreadsheetXml_LeadingWhitespaceBeforeDeclaration_IsTolerated()
    {
        using var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("SpreadsheetXml2003_LeadingWhitespace.xml"));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("ok"));
    }

    [Test]
    public void ReadSpreadsheetXml_ScanMode_UsesActualRowAndColumnCounts()
    {
        using var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("SpreadsheetXml2003_WrongExpandedCounts.xml"));

        Assert.That(reader.FieldCount, Is.EqualTo(2));
        Assert.That(reader.RowCount, Is.EqualTo(2));
    }

    [Test]
    public void ReadSpreadsheetXml_SinglePassMode_HonorsRowSpan()
    {
        using var reader = ExcelReaderFactory.CreateReader(
            Configuration.GetTestWorkbook("SpreadsheetXml2003_RowSpan.xml"),
            new ExcelReaderConfiguration { SinglePassMode = true });

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("First"));

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.IsDBNull(0), Is.True);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.IsDBNull(0), Is.True);

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(0), Is.EqualTo("After span"));

        Assert.That(reader.Read(), Is.False);
    }

    protected override string GetFixtureWorksheetsRowsAndTypes() => "SpreadsheetXml2003";

    protected override string GetFixtureHiddenRow() => "SpreadsheetXml2003_HiddenRefreshRow";

    protected override string GetFixtureSinglePassFieldCount() => "SpreadsheetXml2003_SinglePassFieldCount";

    protected override string GetFixtureVisibility() => "SpreadsheetXml2003_Visibility";

    protected override string GetFixtureMergeCells() => "SpreadsheetXml2003_MergedCells";

    protected override Stream OpenStream(string name)
    {
        return Configuration.GetTestWorkbook(name + ".xml");
    }

    protected override IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null)
    {
        return ExcelReaderFactory.CreateReader(stream, configuration);
    }

    protected override bool IsFixtureSupported(string name) => name switch
    {
        "10x10" => true,
        "10x10000" => true,
        "255x10" => true,
        "Open" => true,
        "MultiSheet" => true,
        "CollapsedHide" => true,
        "BlankHeader" => true,
        "roo_1900_base" => true,
        "roo_1904_base" => true,
        "DoublePrecision" => true,
        "Decimal1109" => true,
        "UnicodeChars" => true,
        "Issue329_Error" => true,
        "OldIssue10725" => true,
        "OldIssue11397" => true,
        "OldIssue11573_BlankValues" => true,
        "OldIssue4031_NullColumn" => true,
        "OldIssue11435_Colors" => true,
        "OldIssue11479_BlankSheet" => true,
        "OldIssue11773_Exponential" => true,
        "OldIssue7433_IllegalOleAutDate" => true,
        "OldIssue8536" => true,
        "DateFormatButNotDate" => true,
        "SpreadsheetXml2003" => true,
        "SpreadsheetXml2003_HiddenRefreshRow" => true,
        "SpreadsheetXml2003_SinglePassFieldCount" => true,
        "SpreadsheetXml2003_Visibility" => true,
        "SpreadsheetXml2003_MergedCells" => true,
        "SpreadsheetXml2003_LeadingWhitespace" => true,
        "SpreadsheetXml2003_WrongExpandedCounts" => true,
        "BoolFormula" => true,
        "Chess" => true,
        "LotsOfSheets" => true,
        "Issue250_Richtext" => true,
        "ColumnWidthsTest" => true,
        "Issue532" => true,
        "Issue574" => true,
        "Issue694_TimeSpan" => true,
        "Issue694_TimeSpanFormula" => true,
        "MergedCell" => true,
        "ExcelDataset" => true,
        "EncodingFormulaDate1520" => true,
        "Issue224_Simple" => true,
        "Issue270" => true,
        "Issue283_TimeSpan" => true,
        "NumDoubleDateBoolString" => true,
        _ => false,
    };
}
