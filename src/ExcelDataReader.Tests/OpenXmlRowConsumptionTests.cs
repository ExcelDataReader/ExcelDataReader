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

    [TestCase(false)]
    [TestCase(true)]
    public void FragmentsResumeExactlyOnceAtEveryCallback(bool reorder)
    {
        string xml = $$"""
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
              <row r="1" ht="24"><c><v>1</v></c></row>
              <row r="1" ht="9"><c r="B1"><v>2</v></c></row>
              <row r="2"><c><v>3</v></c></row>
              <row r="{{(reorder ? 1 : 2)}}"><c><v>4</v></c></row>
              <row r="{{(reorder ? 1 : 2)}}"><c><v>5</v></c></row>
            </sheetData></worksheet>
            """;
        using var parser = new XmlWorksheetReader(XmlReader.Create(new StringReader(xml)), WorksheetReadMode.Materialize, new CommonWorkbook());
        var order = new OpenXmlRowOrder();
        foreach (int index in new[] { 0, 0, 1, reorder ? 0 : 1, reorder ? 0 : 1 })
            order.Add(index);
        using var reader = new OpenXmlRowReader(parser, 15, order.Build());
        List<string> events = [];
        var handler = new PausingHandler(events);
        int pauses = 0;
        while (reader.Consume(ref handler) == WorksheetReadStatus.Paused)
            pauses++;

        string[] expected;
        if (reorder)
        {
            expected =
            [
                "row:0:15", "cell:0:1", "cell:1:2", "cell:0:4", "cell:0:5", "end",
                "row:1:15", "cell:0:3", "end",
            ];
        }
        else
        {
            expected =
            [
                "row:0:24", "cell:0:1", "row:0:9", "cell:1:2", "end",
                "row:1:15", "cell:0:3", "row:1:15", "cell:0:4", "row:1:15", "cell:0:5", "end",
            ];
        }

        Assert.That(events, Is.EqualTo(expected));
        Assert.That(pauses, Is.EqualTo(events.Count));
        Assert.That(reader.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.EndOfInput));
    }

    [TestCase(new[] { 0, 0, 1, 2 }, new int[0])]
    [TestCase(new[] { 2, 0 }, new[] { 0, 1 })]
    [TestCase(new[] { 0, 2, 1, 3 }, new[] { 1, 2 })]
    [TestCase(new[] { 0, 1, 0 }, new[] { 0, 2 })]
    [TestCase(new[] { 0, 2, 1, 5, 4, 6 }, new[] { 1, 2, 3, 4 })]
    [TestCase(new[] { 0, 3, 1, 4, 2, 5 }, new[] { 1, 4 })]
    [TestCase(new[] { 0, 3, 1, 4, 0, 5 }, new[] { 0, 4 })]
    [TestCase(new[] { 0, 2, 4, 6, 3, 8 }, new[] { 2, 4 })]
    [TestCase(new[] { 1, 1, 1, 4, 4, 2, 2, 5 }, new[] { 3, 6 })]
    [TestCase(new[] { 0, 100000, 200000, 300000, 150000, 400000 }, new[] { 2, 4 })]
    [TestCase(new[] { 8, 6, 4, 2, 0 }, new[] { 0, 4 })]
    [TestCase(new[] { 0, 9, 6, 3, 12 }, new[] { 1, 3 })]
    [TestCase(new[] { 0, 0, 2, 2, 1, 1, 1, 3 }, new[] { 2, 6 })]
    [TestCase(new[] { 0, 0, 1, 1, 2, 2, 1, 1, 3, 3 }, new[] { 2, 7 })]
    [TestCase(new[] { 0, 0, 3, 3, 2, 2, 1, 1, 4, 4 }, new[] { 2, 7 })]
    public void OrderIndexFindsOnlySafeOverlappingRegions(int[] rows, int[] boundaries)
    {
        var order = new OpenXmlRowOrder();
        foreach (int row in rows)
            order.Add(row);

        Assert.That(order.Build().SelectMany(region => new[] { region.Start, region.End }), Is.EqualTo(boundaries));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void CompressedRunsPreserveOverlapBoundariesForMixedStridesAndFragments(bool repeatedGroups)
    {
        var random = new Random(741);
        for (int sample = 0; sample < 40; sample++)
        {
            List<int> rows = [];
            for (int run = 0; run < 10; run++)
            {
                int first = random.Next(500);
                int step = random.Next(-5, 6);
                int count = random.Next(1, 16);
                for (int i = 0; i < count; i++)
                {
                    int fragments = repeatedGroups ? random.Next(1, 5) : 1;
                    for (int fragment = 0; fragment < fragments; fragment++)
                        rows.Add(first + i * step);
                }
            }

            List<OpenXmlRowOrder.Region> expected = [];
            int start = 0;
            for (int end = 0; end < rows.Count; end++)
            {
                if (end + 1 < rows.Count && rows.Take(end + 1).Max() >= rows.Skip(end + 1).Min())
                    continue;

                var segment = rows.Skip(start).Take(end - start + 1).ToArray();
                if (!segment.SequenceEqual(segment.OrderBy(index => index)))
                    expected.Add(new(start, end));
                start = end + 1;
            }

            var order = new OpenXmlRowOrder();
            foreach (int row in rows)
                order.Add(row);

            Assert.That(order.Build(), Is.EqualTo(expected), $"sample {sample}");
        }
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