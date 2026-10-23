#nullable enable

using System.Globalization;
using System.Text;
using ExcelDataReader.Core;
using ExcelDataReader.Core.XmlSpreadsheetFormat;

namespace ExcelDataReader.Tests;

[TestFixture]
public sealed class SpreadsheetXmlConsumptionTests
{
    private enum CallbackKind
    {
        None,
        Row,
        Cell,
        RowEnd,
    }

    [TestCase(false)]
    [TestCase(true)]
    public void RealFixtureScanAndTraversalTerminate(bool singlePass)
    {
        using var workbook = new SpreadsheetXmlWorkbook(Configuration.GetTestWorkbook("SpreadsheetXml2003.xml"))
        {
            SinglePassMode = singlePass,
        };
        int sheetIndex = 0;
        foreach (var worksheet in workbook.ReadWorksheets())
        {
            var state = new CallbackState();
            var handler = new RecordingHandler(state);
            using var driver = worksheet.OpenReader();
            Assert.That(driver.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.EndOfInput));
            Assert.That(driver.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.EndOfInput));
            Assert.That(state.Events.Count(value => value == "row-end"), Is.EqualTo(sheetIndex == 0 ? 3 : 1));
            if (!singlePass)
            {
                Assert.That(worksheet.RowCount, Is.EqualTo(sheetIndex == 0 ? 4 : 1));
                Assert.That(worksheet.FieldCount, Is.EqualTo(sheetIndex == 0 ? 4 : 1));
            }

            sheetIndex++;
        }

        Assert.That(sheetIndex, Is.EqualTo(2));
    }

    [Test]
    public void PausingAtRowBeginDoesNotConsumeCellsAndResumesExactlyOnce()
    {
        using var reader = CreateReader(
            """
            <Worksheet ss:Name="Push"><Table>
              <Row ss:Index="2"><Cell><Data ss:Type="String">value</Data></Cell></Row>
            </Table></Worksheet>
            """);
        var state = new CallbackState { PauseAt = CallbackKind.Row };
        var handler = new RecordingHandler(state);

        Assert.That(reader.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.Paused));
        Assert.That(state.Events, Does.Contain("row:1"));
        Assert.That(state.Events.Any(value => value.StartsWith("cell:", StringComparison.Ordinal)), Is.False);

        while (reader.Consume(ref handler) == WorksheetReadStatus.Paused)
        {
        }

        Assert.That(state.Events.Count(value => value == "row:1"), Is.EqualTo(1));
        Assert.That(state.Events.Count(value => value == "cell:0:value"), Is.EqualTo(1));
        Assert.That(state.Events.Count(value => value == "row-end"), Is.EqualTo(1));
    }

    [Test]
    public void PausingAtCellAndRowEndPreservesCompletedRow()
    {
        using var reader = CreateReader(
            """
            <Worksheet ss:Name="Push"><Table>
              <Row><Cell ss:MergeAcross="1"><Data ss:Type="String">first<![CDATA[ second]]></Data></Cell></Row>
              <Row><Cell><Data ss:Type="Number">2.5</Data></Cell></Row>
            </Table></Worksheet>
            """);
        var state = new CallbackState { PauseAt = CallbackKind.Cell };
        var handler = new RecordingHandler(state);

        Assert.That(reader.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.Paused));
        Assert.That(state.Events, Does.Contain("cell:0:first second"));

        state.PauseAt = CallbackKind.RowEnd;
        Assert.That(reader.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.Paused));
        Assert.That(state.Events.Count(value => value == "row-end"), Is.EqualTo(1));

        while (reader.Consume(ref handler) == WorksheetReadStatus.Paused)
        {
        }

        Assert.That(state.Events.Count(value => value == "cell:0:first second"), Is.EqualTo(1));
        Assert.That(state.Events.Count(value => value == "row-end"), Is.EqualTo(2));
        Assert.That(state.Events.Count(value => value == "row:0"), Is.EqualTo(1));
        Assert.That(state.Events.Count(value => value == "row:1"), Is.EqualTo(1));
        Assert.That(state.Events, Does.Contain("merge:0:0:1:0"));
        Assert.That(state.Events, Does.Contain("cell:0:2.5"));
    }

    private static IWorksheetReader CreateReader(string worksheet)
    {
        string xml =
            "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" " +
            "xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">" +
            worksheet +
            "</Workbook>";
        return SpreadsheetXmlWorksheet.Create(
            new MemoryStream(Encoding.UTF8.GetBytes(xml)),
            0,
            "Push",
            "visible",
            null,
            null,
            0,
            new Dictionary<string, ExtendedFormat> { ["Default"] = ExtendedFormat.Zero },
            null,
            singlePassMode: true).OpenReader();
    }

    private struct RecordingHandler(CallbackState state) : IWorksheetHandler
    {
        public WorksheetReadControl OnDimension(CellRange range) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnCodeName(string? codeName) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnDefaultRowHeight(double? height) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnColumn(Column column) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnMergeCell(CellRange range)
        {
            state.Events.Add($"merge:{range.FromColumn}:{range.FromRow}:{range.ToColumn}:{range.ToRow}");
            return WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnHeaderFooter(HeaderFooter headerFooter) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnSheetDataBegin() => WorksheetReadControl.Continue;

        public WorksheetReadControl OnScanResult(int maxRowIndex, int maxColumnIndex) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnRow(int index, bool hidden, double? height)
        {
            state.Events.Add($"row:{index}");
            return state.Pause(CallbackKind.Row) ? WorksheetReadControl.Pause : WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnCell(in Cell cell)
        {
            state.Events.Add(string.Format(CultureInfo.InvariantCulture, "cell:{0}:{1}", cell.ColumnIndex, cell.Value));
            return state.Pause(CallbackKind.Cell) ? WorksheetReadControl.Pause : WorksheetReadControl.Continue;
        }

        public WorksheetReadControl OnRowEnd()
        {
            state.Events.Add("row-end");
            return state.Pause(CallbackKind.RowEnd) ? WorksheetReadControl.Pause : WorksheetReadControl.Continue;
        }
    }

    private sealed class CallbackState
    {
        public List<string> Events { get; } = [];

        public CallbackKind PauseAt { get; set; }

        public bool Pause(CallbackKind kind)
        {
            if (PauseAt != kind)
                return false;
            PauseAt = CallbackKind.None;
            return true;
        }
    }
}