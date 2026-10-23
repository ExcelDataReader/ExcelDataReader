using System.Xml;
using ExcelDataReader.Core;
using ExcelDataReader.Core.OpenXmlFormat;
using ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

namespace ExcelDataReader.Tests;

public class OpenXmlRowConsumptionTests
{
    [TestCase(false)]
    [TestCase(true)]
    public void RepeatedRowFixturePreservesMergedValuesAcrossReset(bool singlePass)
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(
            Configuration.GetTestWorkbook("Issue324.xlsx"),
            new ExcelReaderConfiguration { SinglePassMode = singlePass });
        for (int pass = 0; pass < 2; pass++)
        {
            int rows = 0;
            while (reader.Read())
            {
                if (rows == 10)
                {
                    Assert.That(reader.GetString(2), Is.EqualTo("Other"));
                    Assert.That(reader.GetDouble(3), Is.EqualTo(191036.15));
                    Assert.That(reader.GetDouble(12), Is.EqualTo(67024.6));
                }

                rows++;
            }

            Assert.That(rows, Is.EqualTo(20));
            Assert.That(reader.FieldCount, Is.EqualTo(13));
            reader.Reset();
        }
    }

    [Test]
    public void FragmentAndDescendingHeadersResumeExactlyOnceAtEveryCallback()
    {
        const string xml = """
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
              <row r="1" ht="24"><c><v>1</v></c></row>
              <row r="1" ht="9"><c r="B1"><v>2</v></c></row>
              <row r="2"><c><v>3</v></c></row>
              <row r="1"><c><v>4</v></c></row>
              <row r="1"><c><v>5</v></c></row>
            </sheetData></worksheet>
            """;
        using var parser = new XmlWorksheetReader(XmlReader.Create(new StringReader(xml)), WorksheetReadMode.Materialize, new CommonWorkbook());
        using var reader = new OpenXmlRowReader(parser, 15);
        List<string> events = [];
        var handler = new PausingHandler(events);
        int pauses = 0;
        while (reader.Consume(ref handler) == WorksheetReadStatus.Paused)
            pauses++;

        Assert.That(events, Is.EqualTo(new[]
        {
            "row:0:24", "cell:0:1", "row:0:9", "cell:1:2", "end",
            "row:1:15", "cell:0:3", "end",
            "row:2:15", "cell:0:4", "end",
            "row:3:15", "cell:0:5", "end",
        }));
        Assert.That(pauses, Is.EqualTo(events.Count));
        Assert.That(reader.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.EndOfInput));
    }

    private struct PausingHandler(List<string> events) : IWorksheetHandler
    {
        public WorksheetReadControl OnDimension(CellRange range) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnCodeName(string codeName) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnDefaultRowHeight(double? height) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnColumn(Column column) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnMergeCell(CellRange range) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnHeaderFooter(HeaderFooter headerFooter) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnSheetDataBegin() => WorksheetReadControl.Continue;

        public WorksheetReadControl OnScanResult(int maxRowIndex, int maxColumnIndex) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnRow(int index, bool hidden, double? height) => Add(FormattableString.Invariant($"row:{index}:{height}"));

        public WorksheetReadControl OnCell(in Cell cell) => Add(FormattableString.Invariant($"cell:{cell.ColumnIndex}:{cell.Value}"));

        public WorksheetReadControl OnRowEnd() => Add("end");

        private WorksheetReadControl Add(string value)
        {
            events.Add(value);
            return WorksheetReadControl.Pause;
        }
    }
}
