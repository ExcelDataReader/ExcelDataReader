namespace ExcelDataReader.Tests;

public abstract class ExcelWorkbookTestBase : ExcelTestBase
{
    [TestCase(false)]
    [TestCase(true)]
    public void Issue788_PreservesDateAndDurationPrecision(bool singlePass)
    {
        using var reader = OpenReader(OpenFixtureStream("Issue788_Precision"), new ExcelReaderConfiguration { SinglePassMode = singlePass });
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetValue(1), Is.EqualTo(new DateTime(1899, 12, 31).AddTicks(4000)));
            Assert.That(reader.GetValue(2), Is.EqualTo(TimeSpan.FromSeconds(11)));
            Assert.That(reader.GetValue(3), Is.EqualTo(TimeSpan.FromTicks(-4000)));
            Assert.That(reader.GetValue(4), Is.EqualTo(new DateTime(1899, 12, 31).AddTicks(3219)));
            Assert.That(reader.GetValue(5), Is.EqualTo(0.0004 / 86400));
            Assert.That(reader.Read(), Is.False);
            reader.Reset();
        }
    }

    [Test]
    public void Issue245_CodeName()
    {
        using var reader = OpenReader("ExcelDataset");
        Assert.That(reader.CodeName, Is.EqualTo("Sheet1"));
    }

    [Test]
    public void Issue541_BuiltinFormat55IsDate()
    {
        using var reader = OpenReader("Issue541");
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetValue(0), Is.EqualTo(new DateTime(2021, 1, 15)));
    }
}
