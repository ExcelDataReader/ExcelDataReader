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
    public void GitIssue541BuiltinFormat55IsDate()
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(Configuration.GetTestWorkbook("Test_git_issue_541.xlsb"));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetValue(0), Is.EqualTo(new DateTime(2021, 1, 15)));
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
