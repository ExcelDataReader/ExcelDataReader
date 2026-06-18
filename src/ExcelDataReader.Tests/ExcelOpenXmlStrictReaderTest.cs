using System.Data;

namespace ExcelDataReader.Tests;

public class ExcelOpenXmlStrictReaderTest : ExcelOpenXmlReaderBase
{
    protected override DateTime Issue82_TodayDate => new(2013, 4, 19);

    [TestCase("Issue498")]
    public void Issue498_ReadStrictOpenXmlExcelFile(string fileName)
    {
        using IExcelDataReader reader = OpenReader(fileName);
        DataTableCollection tables = reader.AsDataSet().Tables;

        Assert.That(tables.Count, Is.EqualTo(2));

        foreach (DataTable table in tables)
        {
            Assert.That(table.Rows.Count, Is.EqualTo(2));
            Assert.That(table.Columns.Count, Is.EqualTo(2));
            Assert.That(table.Rows[0][0].ToString(), Is.EqualTo("A1"));
        }
    }

    [Test]
    public void Issue734_StrictDateTimeCellsReturnDateTime()
    {
        using IExcelDataReader reader = OpenReader("NumDoubleDateBoolString");
        Assert.That(reader.Read(), Is.True);

        var value = reader.GetValue(5);
        Assert.That(value, Is.TypeOf<DateTime>());
        Assert.That(reader.GetFieldType(5), Is.EqualTo(typeof(DateTime)));
        Assert.That(reader.GetDateTime(5), Is.EqualTo((DateTime)value));
    }

    protected override IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null)
    {
        return ExcelReaderFactory.CreateOpenXmlReader(stream, configuration);
    }

    protected override Stream OpenStream(string name)
    {
        return Configuration.GetTestWorkbook(Path.Combine("strict", name + ".xlsx"));
    }
}