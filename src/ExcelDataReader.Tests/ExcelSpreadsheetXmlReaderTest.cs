namespace ExcelDataReader.Tests;

[TestFixture]
public class ExcelSpreadsheetXmlReaderTest : ExcelSpreadsheetContractTestBase
{
    protected override DateTime Issue82_TodayDate => new(2013, 4, 19);

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
        "Open" => true,
        "MultiSheet" => true,
        "CollapsedHide" => true,
        "BlankHeader" => true,
        "roo_1900_base" => true,
        "roo_1904_base" => true,
        "SpreadsheetXml2003" => true,
        "SpreadsheetXml2003_HiddenRefreshRow" => true,
        "SpreadsheetXml2003_SinglePassFieldCount" => true,
        "SpreadsheetXml2003_Visibility" => true,
        "SpreadsheetXml2003_MergedCells" => true,
        "SpreadsheetXml2003_LeadingWhitespace" => true,
        "SpreadsheetXml2003_WrongExpandedCounts" => true,
        _ => false,
    };
}
