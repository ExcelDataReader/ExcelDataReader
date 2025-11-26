namespace ExcelDataReader.Tests;

[TestFixture]
public class ExcelOpenXmlBinaryReaderTest : ExcelOpenXmlReaderBase
{
    /// <inheritdoc />
    protected override DateTime GitIssue82TodayDate => new(2013, 4, 19);

    [Test]
    public void GitIssue635()
    {
        using var reader = OpenReader("Test_git_issue_635");
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].Rows[0].ItemArray, Is.EqualTo(new[] { "A", "B", "C", "D", "E", "F" }));
    }

    [Test]
    public void GitIssue642_ActiveSheet()
    {
        using var reader = OpenReader("Test_git_issue_642");
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            FilterSheet = (tableReader, sheetIndex) => tableReader.IsActiveSheet
        });
        Assert.That(reader.ActiveSheet, Is.EqualTo(5));
        Assert.That(dataSet.Tables[0].TableName, Is.EqualTo("List6"));
    }

    [Test]
    public void GitIssue642_ActiveSheet_SingleWorksheet()
    {
        using var reader = OpenReader("Test_git_issue_642onesheet");
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            FilterSheet = (tableReader, sheetIndex) => tableReader.IsActiveSheet
        });
        Assert.That(reader.ActiveSheet, Is.EqualTo(0));
        Assert.That(dataSet.Tables[0].TableName, Is.EqualTo("List1"));
    }

    [Test]
    public void ClipboardDimension()
    {
        using var excelReader = OpenReader("Test_Clipboard_Biff12");

        Assert.That(excelReader.Dimension.FromRow, Is.EqualTo(10));
        Assert.That(excelReader.Dimension.ToRow, Is.EqualTo(11));
        Assert.That(excelReader.Dimension.FromColumn, Is.EqualTo(6));
        Assert.That(excelReader.Dimension.ToColumn, Is.EqualTo(7));
    }

    /// <inheritdoc />
    protected override Stream OpenStream(string name)
    {
        return Configuration.GetTestWorkbook(name + ".xlsb");
    }

    /// <inheritdoc />
    protected override IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null)
    {
        return ExcelReaderFactory.CreateOpenXmlReader(stream, configuration);
    }
}
