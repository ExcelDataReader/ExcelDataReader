using System.Data;
using System.IO.Compression;
using System.Text;

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

    [Test]
    public void Issue759_TimeOnlyCellsUseExcelBaseDate()
    {
        using IExcelDataReader reader = OpenReader("LocaleTime");
        var dataSet = reader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[1][1], Is.EqualTo(new DateTime(1899, 12, 31, 1, 34, 0)));
        Assert.That(dataSet.Tables[0].Rows[2][1], Is.EqualTo(new DateTime(1899, 12, 31, 1, 34, 0)));
        Assert.That(dataSet.Tables[0].Rows[3][1], Is.EqualTo(new DateTime(1899, 12, 31, 18, 47, 0)));
    }

    [Test]
    public void Issue759_TimeOnlyCellsUse1904BaseDateWithoutChangingExplicitDates()
    {
        using var workbook = new MemoryStream();
        using (var source = OpenStream("LocaleTime"))
            source.CopyTo(workbook);

        using (var archive = new ZipArchive(workbook, ZipArchiveMode.Update, leaveOpen: true))
        {
            ReplaceEntry(archive, "xl/workbook.xml", xml => xml.Replace("<workbookPr ", "<workbookPr date1904=\"1\" "));
            ReplaceEntry(archive, "xl/worksheets/sheet1.xml", xml => xml.Replace("<v>2012-11-28</v>", "<v>0001-01-01</v>"));
        }

        workbook.Position = 0;
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(workbook);
        var dataSet = reader.AsDataSet();

        Assert.That(dataSet.Tables[0].Rows[1][0], Is.EqualTo(new DateTime(1, 1, 1)));
        Assert.That(dataSet.Tables[0].Rows[1][1], Is.EqualTo(new DateTime(1904, 1, 1, 1, 34, 0)));
    }

    protected override IExcelDataReader OpenReader(Stream stream, ExcelReaderConfiguration configuration = null)
    {
        return ExcelReaderFactory.CreateOpenXmlReader(stream, configuration);
    }

    protected override Stream OpenStream(string name)
    {
        return Configuration.GetTestWorkbook(Path.Combine("strict", name + ".xlsx"));
    }

    private static void ReplaceEntry(ZipArchive archive, string path, Func<string, string> replace)
    {
        var entry = archive.GetEntry(path)!;
        string xml;
        using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
            xml = reader.ReadToEnd();

        entry.Delete();
        using var writer = new StreamWriter(archive.CreateEntry(path).Open(), new UTF8Encoding(false));
        writer.Write(replace(xml));
    }
}
