using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;
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

    [TestCase(false)]
    [TestCase(true)]
    public void IndexedRowsAndSpans_AreEmittedLazilyWithTheirHeights(bool singlePass)
    {
        using var reader = CreateReader(
            """
            <Worksheet ss:Name="Sparse"><Table ss:ExpandedRowCount="1" ss:ExpandedColumnCount="1">
              <Row ss:Index="3" ss:Span="2" ss:Height="25">
                <Cell ss:Index="3"><Data ss:Type="String">third</Data></Cell>
              </Row>
              <Row ss:Index="7" ss:Hidden="1"><Cell><Data ss:Type="String">hidden</Data></Cell></Row>
              <Row/>
            </Table></Worksheet>
            """,
            singlePass);
        if (!singlePass)
        {
            Assert.That(reader.RowCount, Is.EqualTo(8));
            Assert.That(reader.FieldCount, Is.EqualTo(3));
        }

        double[] heights = [15, 15, 25, 25, 25, 15, 0, 15];
        for (int i = 0; i < heights.Length; i++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.RowHeight, Is.EqualTo(heights[i]));
            if (i == 2)
            {
                Assert.That(reader.GetString(2), Is.EqualTo("third"));
            }
            else if (i == 6)
            {
                Assert.That(reader.GetString(0), Is.EqualTo("hidden"));
            }
            else
            {
                for (int column = 0; column < reader.FieldCount; column++)
                    Assert.That(reader.IsDBNull(column), Is.True);
            }
        }

        Assert.That(reader.Read(), Is.False);
    }

    [Test]
    public void ScanCounts_DataAndMergesButNotEmptyTrailingCells()
    {
        using var reader = CreateReader(
            """
            <Worksheet ss:Name="Counts"><Table ss:ExpandedRowCount="99" ss:ExpandedColumnCount="99">
              <Column ss:Index="2" ss:Span="1" ss:Hidden="1"/>
              <Row><Cell ss:Index="2"><Data ss:Type="String"/></Cell><Cell ss:Index="99"/></Row>
              <Row><Cell ss:Index="4" ss:MergeAcross="2" ss:MergeDown="1"/></Row>
              <Row/>
            </Table></Worksheet>
            """);
        Assert.That(reader.RowCount, Is.EqualTo(3));
        Assert.That(reader.FieldCount, Is.EqualTo(6));
        Assert.That(reader.GetColumnWidth(1), Is.Zero);
        Assert.That(reader.GetColumnWidth(2), Is.Zero);
        Assert.That(reader.MergeCells, Has.Length.EqualTo(1));
        Assert.That(reader.MergeCells[0].FromColumn, Is.EqualTo(3));
        Assert.That(reader.MergeCells[0].ToColumn, Is.EqualTo(5));
        Assert.That(reader.MergeCells[0].ToRow, Is.EqualTo(2));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(1), Is.Empty);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.IsDBNull(3), Is.True);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.Read(), Is.False);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void ResetAndNextResult_ReopenStreamsAfterPartialAndCompleteReads(bool singlePass)
    {
        using var reader = CreateReader(
            """
            <Worksheet ss:Name="First"><Table>
              <Row><Cell><Data ss:Type="String">first</Data></Cell></Row>
              <Row><Cell><Data ss:Type="String">second</Data></Cell></Row>
            </Table></Worksheet>
            <Worksheet ss:Name="Empty"><Table/></Worksheet>
            <Worksheet ss:Name="Last"><Table>
              <Row><Cell><Data ss:Type="Number">42</Data></Cell></Row>
            </Table></Worksheet>
            """,
            singlePass);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.NextResult(), Is.True);
        Assert.That(reader.Name, Is.EqualTo("Empty"));
        if (!singlePass)
        {
            Assert.That(reader.RowCount, Is.Zero);
            Assert.That(reader.FieldCount, Is.Zero);
        }

        Assert.That(reader.Read(), Is.False);
        Assert.That(reader.NextResult(), Is.True);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(42));
        reader.Reset();
        for (int traversal = 0; traversal < 2; traversal++)
        {
            Assert.That(reader.Name, Is.EqualTo("First"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo("first"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo("second"));
            Assert.That(reader.Read(), Is.False);
            Assert.That(reader.NextResult(), Is.True);
            Assert.That(reader.Read(), Is.False);
            Assert.That(reader.NextResult(), Is.True);
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(0), Is.EqualTo(42));
            Assert.That(reader.Read(), Is.False);
            Assert.That(reader.NextResult(), Is.False);
            reader.Reset();
        }
    }

    [TestCase(false)]
    [TestCase(true)]
    public void NonAscendingRows_FailExplicitly(bool singlePass)
    {
        using var reader = CreateReader(
            """
            <Worksheet ss:Name="Invalid"><Table><Row/><Row ss:Index="1"/></Table></Worksheet>
            """,
            singlePass);
        Assert.That(reader.Read(), Is.True);
        Assert.Throws<XmlException>(() => reader.Read());
    }

    [Test]
    [NonParallelizable]
    public void ScanMode_RetainedMemoryDoesNotGrowWithCellPayloadOrVisitedSheets()
    {
        MeasureRetainedMemory(10);
        long small = MeasureRetainedMemory(100);
        long large = MeasureRetainedMemory(10000);
        TestContext.Out.WriteLine($"Retained bytes: 100 rows/sheet = {small}; 10000 rows/sheet = {large}");
        Assert.That(large - small, Is.LessThan(8L * 1024 * 1024));
    }

    [Test]
    [NonParallelizable]
    public void ScanMode_LargeRowIndexDoesNotMaterializeEmptyRows()
    {
        long before = GC.GetTotalMemory(true);
        using var reader = CreateReader(
            """
            <Worksheet ss:Name="Sparse"><Table><Row ss:Index="1000000">
              <Cell><Data ss:Type="String">last</Data></Cell>
            </Row></Table></Worksheet>
            """);
        Assert.That(reader.RowCount, Is.EqualTo(1000000));
        Assert.That(reader.FieldCount, Is.EqualTo(1));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.IsDBNull(0), Is.True);
        long retained = GC.GetTotalMemory(true) - before;
        GC.KeepAlive(reader);
        Assert.That(retained, Is.LessThan(8L * 1024 * 1024));
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

    private static IExcelDataReader CreateReader(string worksheets, bool singlePass = false)
    {
        string xml = "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">" +
            worksheets + "</Workbook>";
        return ExcelReaderFactory.CreateReader(
            new MemoryStream(Encoding.UTF8.GetBytes(xml)),
            new ExcelReaderConfiguration { SinglePassMode = singlePass });
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static long MeasureRetainedMemory(int rowsPerSheet)
    {
        using var stream = new FileStream(Path.GetTempFileName(), FileMode.Open, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose);
        using (var writer = new StreamWriter(stream, new UTF8Encoding(false), 4096, leaveOpen: true))
        {
            writer.Write("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");
            string padding = new('a', 4096);
            for (int sheet = 0; sheet < 2; sheet++)
            {
                writer.Write("<Worksheet ss:Name=\"Sheet");
                writer.Write(sheet.ToString(CultureInfo.InvariantCulture));
                writer.Write("\"><Table>");
                for (int row = 0; row < rowsPerSheet; row++)
                {
                    writer.Write("<Row><Cell><Data ss:Type=\"String\">");
                    writer.Write(row.ToString(CultureInfo.InvariantCulture));
                    writer.Write(padding);
                    writer.Write("</Data></Cell></Row>");
                }

                writer.Write("</Table></Worksheet>");
            }

            writer.Write("</Workbook>");
        }

        stream.Position = 0;
        long before = GC.GetTotalMemory(true);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        long maximum = GC.GetTotalMemory(true) - before;
        do
        {
            Assert.That(reader.RowCount, Is.EqualTo(rowsPerSheet));
            for (int row = 0; row < rowsPerSheet; row++)
            {
                Assert.That(reader.Read(), Is.True);
                Assert.That(reader.GetString(0), Does.StartWith(row.ToString(CultureInfo.InvariantCulture)));
                if (row == rowsPerSheet / 2)
                    maximum = Math.Max(maximum, GC.GetTotalMemory(true) - before);
            }

            Assert.That(reader.Read(), Is.False);
            maximum = Math.Max(maximum, GC.GetTotalMemory(true) - before);
        }
        while (reader.NextResult());
        GC.KeepAlive(reader);
        return maximum;
    }
}
