namespace ExcelDataReader.Tests;

[TestFixture]
public class ExcelOpenXmlBinaryReaderTest : ExcelOpenXmlReaderBase
{
    /// <inheritdoc />
    protected override DateTime Issue82_TodayDate => new(2013, 4, 19);

    protected override bool UsesBinaryWorksheet => true;

    [Test]
    public void Issue635()
    {
        using var reader = OpenReader("Issue635");
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].Rows[0].ItemArray, Is.EqualTo(new[] { "A", "B", "C", "D", "E", "F" }));
    }

    [Test]
    public void ClipboardDimension()
    {
        using var excelReader = OpenReader("ClipboardBiff12");

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