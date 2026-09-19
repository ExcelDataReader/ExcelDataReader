namespace ExcelDataReader.Tests;

public abstract class ExcelSpreadsheetContractTestBase : ExcelTestBase
{
    [Test]
    public void SpreadsheetContract_WorksheetsRowsAndTypes()
    {
        using var reader = OpenReader(OpenFixtureStream(GetFixtureWorksheetsRowsAndTypes()));

        Assert.That(reader.ResultsCount, Is.EqualTo(2));
        Assert.That(reader.Name, Is.EqualTo("Sheet1"));
        Assert.That(reader.CodeName, Is.EqualTo("Sheet1"));
        Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft&CCenter&RRight"));
        Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LFoot&CFooter&RPage &P"));
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
    public void SpreadsheetContract_SinglePassMode_PreservesSheetMetadata()
    {
        using var reader = OpenReader(OpenFixtureStream(GetFixtureWorksheetsRowsAndTypes()), new ExcelReaderConfiguration { SinglePassMode = true });

        Assert.That(reader.Name, Is.EqualTo("Sheet1"));
        Assert.That(reader.CodeName, Is.EqualTo("Sheet1"));
        Assert.That(reader.HeaderFooter?.OddHeader, Is.EqualTo("&LLeft&CCenter&RRight"));
        Assert.That(reader.HeaderFooter?.OddFooter, Is.EqualTo("&LFoot&CFooter&RPage &P"));
    }

    [Test]
    public void SpreadsheetContract_HiddenRow_UsesZeroRowHeight()
    {
        using var reader = OpenReader(OpenFixtureStream(GetFixtureHiddenRow()));

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
    public void SpreadsheetContract_SinglePassMode_RowCountThrowsAndFieldCountGrows()
    {
        using var reader = OpenReader(OpenFixtureStream(GetFixtureSinglePassFieldCount()), new ExcelReaderConfiguration { SinglePassMode = true });

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
    public void SpreadsheetContract_SheetVisibility_AndDataSetProperties()
    {
        using var reader = OpenReader(OpenFixtureStream(GetFixtureVisibility()));
        Assert.That(reader.VisibleState, Is.EqualTo("hidden"));

        Assert.That(reader.NextResult(), Is.True);
        Assert.That(reader.VisibleState, Is.EqualTo("visible"));

        Assert.That(reader.NextResult(), Is.True);
        Assert.That(reader.VisibleState, Is.EqualTo("veryhidden"));

        reader.Reset();
        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables[0].ExtendedProperties["visiblestate"], Is.EqualTo("hidden"));
        Assert.That(dataSet.Tables[1].ExtendedProperties["visiblestate"], Is.EqualTo("visible"));
        Assert.That(dataSet.Tables[2].ExtendedProperties["visiblestate"], Is.EqualTo("veryhidden"));
    }

    [Test]
    public void SpreadsheetContract_MergeCells_AreReported()
    {
        using var reader = OpenReader(OpenFixtureStream(GetFixtureMergeCells()));
        Assert.That(reader.MergeCells, Is.EquivalentTo(new[]
        {
            new[] { 0, 0, 1, 0 },
            new[] { 2, 1, 2, 2 },
        }).Using<CellRange, int[]>((a, e) => a.FromColumn == e[0] && a.FromRow == e[1] && a.ToColumn == e[2] && a.ToRow == e[3]));
    }

    protected abstract string GetFixtureWorksheetsRowsAndTypes();

    protected abstract string GetFixtureHiddenRow();

    protected abstract string GetFixtureSinglePassFieldCount();

    protected abstract string GetFixtureVisibility();

    protected abstract string GetFixtureMergeCells();
}
