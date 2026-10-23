using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Xml;
using ExcelDataReader.Core;
using ExcelDataReader.Core.BinaryFormat;
using ExcelDataReader.Core.OpenXmlFormat;
using ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;
using ExcelDataReader.Core.OpenXmlFormat.XmlFormat;
using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Tests;

public class WorksheetConsumptionTests
{
    [TestCase("xml")]
    [TestCase("xml-preparing")]
    [TestCase("xlsb")]
    [TestCase("xlsb-preparing")]
    public void WorksheetCallbacksPauseAndResumeEveryEventExactlyOnce(string format)
    {
        List<string> events;
        switch (format)
        {
            case "xml":
                events = ReadCallbacks(CreateXmlCallbackReader(preparing: false));
                Assert.That(
                    events,
                    Is.EqualTo(new[]
                {
                    "code", "height", "dimension", "column", "merge", "header", "begin",
                    "row:0", "cell:0:0", "cell:0:1", "row:2", "cell:2:1",
                }),
                    string.Join(", ", events));
                break;
            case "xml-preparing":
                events = ReadCallbacks(CreateXmlCallbackReader(preparing: true));
                Assert.That(events, Is.EqualTo(new[]
                {
                    "code", "height", "dimension", "column", "merge", "header", "begin", "row:0", "row:2", "scan",
                }));
                break;
            case "xlsb":
                events = ReadCallbacks(CreateXlsbCallbackReader(preparing: false));
                Assert.That(events, Is.EqualTo(new[]
                {
                    "dimension", "column", "merge", "begin",
                    "row:0", "cell:0:0", "cell:0:1", "row:2", "cell:2:1",
                }));
                break;
            case "xlsb-preparing":
                events = ReadCallbacks(CreateXlsbCallbackReader(preparing: true));
                Assert.That(events, Is.EqualTo(new[]
                {
                    "dimension", "column", "merge", "begin", "row:0", "row:2", "scan",
                }));
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }
    }

    [TestCase(true)]
    [TestCase(false)]
    public void XlsWorksheetReturnsCompletedRowsInOrder(bool singlePass)
    {
        using var workbook = new XlsWorkbook(new MemoryStream(XlsWorkbook(overlapping: !singlePass)), new ExcelReaderConfiguration());
        workbook.SinglePassMode = singlePass;
        var events = ReadCallbacks(workbook.ReadWorksheets().Single().OpenReader());
        var rowIndexes = events.Where(e => e.StartsWith("row:", StringComparison.Ordinal)).Select(e => int.Parse(e.Substring(4), System.Globalization.CultureInfo.InvariantCulture));
        int cells = events.Count(e => e.StartsWith("cell:", StringComparison.Ordinal));
        Assert.That(rowIndexes, Is.EqualTo(Enumerable.Range(0, 40)));
        Assert.That(cells, Is.EqualTo(41));
    }

    [Test]
    public void XlsTransportFramesPreserveGenericRecordsAndTheirPayloads()
    {
        using var buffer = new MemoryStream();
        using (var writer = new BinaryWriter(buffer, Encoding.Unicode, leaveOpen: true))
        {
            XlsRecord(writer, 0x809, data =>
            {
                data.Write((ushort)0x600);
                data.Write((ushort)0x10);
                data.Write(new byte[12]);
            });
            XlsRecord(writer, 0x00F0, data => data.Write((ushort)123));
            XlsRecord(writer, 0x003D, data =>
            {
                data.Write(new byte[10]);
                data.Write((ushort)1);
            });
            XlsRecord(writer, (ushort)BIFFRECORDTYPE.IXFE, data => data.Write((ushort)7));
            XlsRecord(writer, 0x00F1, data => data.Write((ushort)456));
        }

        using var biffStream = new XlsBiffStream(new MemoryStream(buffer.ToArray()));
        List<BIFFRECORDTYPE> ids = [];
        List<ushort> values = [];
        while (biffStream.ReadFrame(out var frame))
        {
            ids.Add(frame.Id);
            values.Add(frame.ReadUInt16(frame.Id == BIFFRECORDTYPE.WINDOW1 ? 10 : 0));
        }

        Assert.That(ids, Is.EqualTo(new BIFFRECORDTYPE[] { BIFFRECORDTYPE.BOF, (BIFFRECORDTYPE)0x00F0, BIFFRECORDTYPE.WINDOW1, BIFFRECORDTYPE.IXFE, (BIFFRECORDTYPE)0x00F1 }));
        Assert.That(values, Is.EqualTo(new ushort[] { 0x600, 123, 1, 7, 456 }));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void XlsMetadataOwnsStringsUntilTheLaterCodePageAndNextSheetReads(bool singlePass)
    {
        using var buffer = new MemoryStream();
        using (var writer = new BinaryWriter(buffer, Encoding.Unicode, leaveOpen: true))
        {
            WriteBof(5);
            long sheetRecord = buffer.Position;
            XlsRecord(writer, 0x85, data =>
            {
                data.Write(0);
                data.Write((ushort)0);
                data.Write((byte)1);
                data.Write((byte)0xe9);
            });
            XlsRecord(writer, 0x41e, data =>
            {
                data.Write((ushort)164);
                data.Write((byte)1);
                data.Write((byte)0xe9);
            });
            XlsRecord(writer, 0x31, data =>
            {
                data.Write(new byte[12]);
                data.Write((byte)204);
                data.Write((byte)0);
                data.Write((byte)1);
                data.Write((byte)0xe9);
            });
            XlsRecord(writer, 0xe0, data =>
            {
                data.Write((ushort)0);
                data.Write((ushort)164);
                data.Write(new byte[16]);
            });
            XlsRecord(writer, 0x42, data => data.Write((ushort)1251));
            XlsRecord(writer, 0x3d, data =>
            {
                data.Write(new byte[10]);
                data.Write((ushort)2);
            });
            XlsRecord(writer, 0xa, _ => { });
            long sheetStart = buffer.Position;
            buffer.Position = sheetRecord + 4;
            writer.Write((uint)sheetStart);
            buffer.Position = sheetStart;
            WriteBof(0x10);
            foreach (ushort id in new ushort[] { 0x14, 0x15 })
            {
                XlsRecord(writer, id, data =>
                {
                    data.Write((byte)1);
                    data.Write((byte)0xe8);
                });
            }

            XlsRecord(writer, 0x42, data => data.Write((ushort)1251));
            XlsRecord(writer, 0xf0, data => data.Write(new byte[48]));
            XlsRecord(writer, 0x204, data =>
            {
                data.Write((ushort)0);
                data.Write((ushort)0);
                data.Write((ushort)0);
                data.Write((ushort)1);
                data.Write((byte)0xe9);
            });
            XlsRecord(writer, 0xa, _ => { });

            void WriteBof(ushort type) => XlsRecord(writer, 0x809, data =>
            {
                data.Write((ushort)0x500);
                data.Write(type);
                data.Write(new byte[12]);
            });
        }

        using var reader = ExcelReaderFactory.CreateBinaryReader(new MemoryStream(buffer.ToArray()), new ExcelReaderConfiguration { SinglePassMode = singlePass });
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Name, Is.EqualTo("\u0439"));
            Assert.That(reader.ActiveSheet, Is.EqualTo(2));
            Assert.That(reader.HeaderFooter.OddHeader, Is.EqualTo("\u0438"));
            Assert.That(reader.HeaderFooter.OddFooter, Is.EqualTo("\u0438"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo("\u0439"));
            Assert.That(reader.GetNumberFormatString(0), Is.EqualTo("\u0439"));
            Assert.That(reader.Read(), Is.False);
            Assert.That(reader.Name, Is.EqualTo("\u0439"));
            Assert.That(reader.HeaderFooter.OddHeader, Is.EqualTo("\u0438"));
            reader.Reset();
        }
    }

    [TestCase(1)]
    [TestCase(2)]
    [TestCase(3)]
    [TestCase(4)]
    [TestCase(5)]
    [TestCase(6)]
    [TestCase(8)]
    public void XlsSstContinuesSplitHeadersCharactersAndRichTextTailsWithoutRetainingFrames(int headerSplit)
    {
        byte[] header = [3, 0, 0x0c, 2, 0, 3, 0, 0, 0];
        using var buffer = new MemoryStream();
        using (var writer = new BinaryWriter(buffer, Encoding.Unicode, leaveOpen: true))
        {
            XlsRecord(writer, 0x809, data =>
            {
                data.Write((ushort)0x600);
                data.Write((ushort)0x10);
                data.Write(new byte[12]);
            });
            XlsRecord(writer, 0xfc, data =>
            {
                data.Write(2U);
                data.Write(2U);
                data.Write(header, 0, headerSplit);
            });
            XlsRecord(writer, 0x3c, data =>
            {
                data.Write(header, headerSplit, header.Length - headerSplit);
                data.Write((byte)'A');
            });
            XlsRecord(writer, 0x3c, data =>
            {
                data.Write((byte)1);
                data.Write((ushort)'\u03bb');
                data.Write((ushort)'B');
                data.Write(new byte[1]);
            });
            XlsRecord(writer, 0x3c, data =>
            {
                data.Write(new byte[10]);
                data.Write((ushort)4);
                data.Write((byte)0);
                data.Write(Encoding.ASCII.GetBytes("tail"));
            });
            XlsRecord(writer, 0xf0, data => data.Write(new byte[48]));
        }

        var strings = new SharedStringTable();
        var parser = new XlsSSTReader();
        using var transport = new XlsBiffStream(new MemoryStream(buffer.ToArray()));
        while (transport.ReadFrame(out var frame))
        {
            if (frame.Id == BIFFRECORDTYPE.SST)
                parser.ReadStrings(frame, 12, false, strings, 2);
            else if (frame.Id == BIFFRECORDTYPE.CONTINUE)
                parser.ReadStrings(frame, 4, true, strings, 2 - strings.Count);
        }

        parser.Flush(strings);
        Assert.That(strings, Is.EqualTo(new[] { "A\u03bbB", "tail" }));
    }

    [TestCase(2, false)]
    [TestCase(3, false)]
    [TestCase(5, true)]
    public void XlsLegacyFormulaStringContinuesRawCharacterBytesAndPreservesLookahead(int version, bool unicode)
    {
        Encoding encoding = unicode ? Encoding.Unicode : Encoding.GetEncoding(1251);
        byte[] characters = encoding.GetBytes("A\u0439BC");
        int split = characters.Length / 2 - 1;
        using var buffer = new MemoryStream();
        using (var writer = new BinaryWriter(buffer, Encoding.Unicode, leaveOpen: true))
        {
            XlsRecord(writer, 0x809, data =>
            {
                data.Write((ushort)(version == 5 ? 0x500 : version * 0x100));
                data.Write((ushort)0x10);
                data.Write(new byte[12]);
            });
            XlsRecord(writer, version == 2 ? (ushort)7 : (ushort)0x207, data =>
            {
                if (version == 2)
                    data.Write((byte)4);
                else
                    data.Write((ushort)4);
                data.Write(characters, 0, split);
            });
            XlsRecord(writer, 0x3c, data => data.Write(characters, split, characters.Length - split));
            XlsRecord(writer, 0xf0, data => data.Write((ushort)123));
        }

        using var transport = new XlsBiffStream(new MemoryStream(buffer.ToArray()));
        Assert.That(transport.ReadFrame(out _), Is.True);
        string value = XlsFormulaStringReader.Read(transport, version, encoding);
        Assert.That(transport.ReadFrame(out var next), Is.True);
        Assert.That(next.Id, Is.EqualTo((BIFFRECORDTYPE)0xf0));
        Assert.That(next.ReadUInt16(0), Is.EqualTo(123));
        Assert.That(value, Is.EqualTo("A\u0439BC"));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void XlsFormulaStringContinuesWithEncodingChangesAndLeavesTheNextCellReadable(bool singlePass)
    {
        using var buffer = new MemoryStream();
        using (var writer = new BinaryWriter(buffer, Encoding.Unicode, leaveOpen: true))
        {
            XlsRecord(writer, 0x809, data =>
            {
                data.Write((ushort)0x600);
                data.Write((ushort)0x10);
                data.Write(new byte[12]);
            });
            XlsRecord(writer, 6, data =>
            {
                data.Write(new byte[12]);
                data.Write((ushort)0xffff);
                data.Write(new byte[8]);
            });
            XlsRecord(writer, 0x4bc, data => data.Write(new byte[10]));
            XlsRecord(writer, 0x207, data =>
            {
                data.Write((ushort)4);
                data.Write((byte)0);
                data.Write((byte)'A');
            });
            XlsRecord(writer, 0x3c, data =>
            {
                data.Write((byte)1);
                data.Write((ushort)'\u03bb');
            });
            XlsRecord(writer, 0x3c, data =>
            {
                data.Write((byte)0);
                data.Write((byte)'B');
                data.Write((byte)'C');
            });
            XlsRecord(writer, 0x203, data =>
            {
                data.Write((ushort)1);
                data.Write((ushort)0);
                data.Write((ushort)0);
                data.Write(7.0);
            });
            XlsRecord(writer, 0xa, _ => { });
        }

        using var reader = ExcelReaderFactory.CreateBinaryReader(new MemoryStream(buffer.ToArray()), new ExcelReaderConfiguration { SinglePassMode = singlePass });
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Read(), Is.True);
            string value = reader.GetString(0);
            Assert.That(value, Is.EqualTo("A\u03bbBC"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(0), Is.EqualTo(7));
            Assert.That(value, Is.EqualTo("A\u03bbBC"));
            Assert.That(reader.Read(), Is.False);
            reader.Reset();
        }
    }

    [TestCase("xml")]
    [TestCase("xlsb")]
    public void WorksheetCallbackExceptionsPropagateAndReaderCanBeDisposed(string format)
    {
        using IWorksheetReader reader = format == "xml"
            ? CreateXmlCallbackReader(preparing: false)
            : CreateXlsbCallbackReader(preparing: false);
        var handler = default(ThrowingCallbackHandler);

        Assert.Throws<InvalidOperationException>(() => reader.Consume(ref handler));
        Assert.DoesNotThrow(reader.Dispose);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void XlsWorksheetTransportBorrowsAllRecordsThroughOneFrame(bool malformedPackedRange)
    {
        var type = typeof(ExcelReaderFactory).Assembly.GetType("ExcelDataReader.Core.BinaryFormat.XlsBiffStream")!;
        byte[] bytes = XlsWorkbook(false);
        if (malformedPackedRange)
        {
            for (int offset = 0; offset < bytes.Length;)
            {
                ushort id = BitConverter.ToUInt16(bytes, offset);
                int size = BitConverter.ToUInt16(bytes, offset + 2);
                if (id == 0xbd)
                    bytes[offset + 4 + size - 2] = 42;
                offset += 4 + size;
            }
        }

        using var source = new MemoryStream(bytes);
        using var transport = (IDisposable)type.GetConstructors().Single().Invoke([source, 0]);
        var read = type.GetMethod("ReadFrame")!;
        var bufferField = type.GetField("_frameBuffer", BindingFlags.Instance | BindingFlags.NonPublic)!;
        var getBuffer = bufferField.FieldType.GetMethod("GetBuffer")!;
        object[] arguments = [null];
        byte[] borrowed = null;
        int cellRecords = 0;
        int packedRecords = 0;
        ushort packedLastColumn = 0;
        int recordCount = 0;
        while ((bool)read.Invoke(transport, arguments)!)
        {
            object result = arguments[0]!;
            Type resultType = result.GetType();
            Assert.That(resultType.Name, Is.EqualTo("XlsBiffFrame"));
            Assert.That(resultType.GetProperty("OwnedRecord"), Is.Null);
            bool isCell = (bool)resultType.GetProperty("IsCell")!.GetValue(result)!;
            recordCount++;
            var current = (byte[])getBuffer.Invoke(bufferField.GetValue(transport), [1])!;
            if (borrowed != null)
            {
                Assert.That(current, Is.SameAs(borrowed));
            }

            borrowed = current;
            if (isCell)
                cellRecords++;

            if ((ushort)resultType.GetProperty("Id")!.GetValue(result)! == 0x00BD)
            {
                packedRecords++;
                packedLastColumn = (ushort)resultType.GetProperty("LastColumnIndex")!.GetValue(result)!;
            }
        }

        Assert.That(recordCount, Is.GreaterThan(cellRecords));
        Assert.That(cellRecords, Is.EqualTo(40));
        Assert.That(packedRecords, Is.EqualTo(1));
        Assert.That(packedLastColumn, Is.EqualTo(malformedPackedRange ? 42 : 1));
        Assert.That(read.Invoke(transport, arguments), Is.False);
        Assert.That(source.CanRead, Is.True);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void XlsTruncatedNextCellPreservesDecodeFailureAndLeaveOpen(bool shortPayload)
    {
        using var buffer = new MemoryStream();
        using (var writer = new BinaryWriter(buffer, Encoding.Unicode, leaveOpen: true))
        {
            XlsRecord(writer, 0x809, data =>
            {
                data.Write((ushort)0x600);
                data.Write((ushort)0x10);
                data.Write(new byte[12]);
            });
            for (int row = 0; row < 2; row++)
            {
                XlsRecord(writer, 0x208, data =>
                {
                    data.Write((ushort)row);
                    data.Write((ushort)0);
                    data.Write((ushort)1);
                    data.Write((ushort)255);
                    data.Write(0);
                    data.Write(0);
                });
            }

            for (int row = 0; row < 2; row++)
            {
                XlsRecord(writer, 0x203, data =>
                {
                    data.Write((ushort)row);
                    data.Write((ushort)0);
                    data.Write((ushort)0);
                    if (row == 1 && shortPayload)
                        data.Write(1);
                    else
                        data.Write(row + 0.25);
                });
            }

            XlsRecord(writer, 0xa, _ => { });
        }

        byte[] bytes = buffer.ToArray();
        if (!shortPayload)
            Array.Resize(ref bytes, bytes.Length - 10);
        using var source = new MemoryStream(bytes);
        using (var reader = ExcelReaderFactory.CreateReader(source, new ExcelReaderConfiguration { SinglePassMode = true, LeaveOpen = true }))
        {
            if (shortPayload)
            {
                Assert.That(reader.Read(), Is.True);
                Assert.That(reader.GetDouble(0), Is.EqualTo(0.25));
                Assert.Throws<ArgumentException>(() => reader.Read());
                reader.Reset();
                Assert.That(reader.Read(), Is.True);
                Assert.That(reader.GetDouble(0), Is.EqualTo(0.25));
            }
            else
            {
                Assert.Throws<EndOfStreamException>(() => reader.Read());
            }
        }

        Assert.That(source.CanRead, Is.True);
    }

    [TestCase(false, false)]
    [TestCase(true, false)]
    [TestCase(false, true)]
    public void XlsCellConsumptionKeepsPackedCellsFormulaStringsErrorsAndOverlappingRows(bool singlePass, bool overlapping)
    {
        byte[] data = XlsWorkbook(overlapping);
        using var stream = new MemoryStream(data);
        using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration { SinglePassMode = singlePass, LeaveOpen = true }))
        {
            Assert.That(reader.Read(), Is.True);
            reader.Reset();
            for (int pass = 0; pass < 2; pass++)
            {
                string owned = null;
                for (int row = 0; row < 40; row++)
                {
                    Assert.That(reader.Read(), Is.True);
                    if (row == 1)
                    {
                        Assert.That(reader.GetDouble(0), Is.EqualTo(7));
                        Assert.That(reader.GetDouble(1), Is.EqualTo(8));
                        Assert.That(reader.RowHeight, Is.EqualTo(24));
                    }
                    else if (row == 2)
                    {
                        owned = reader.GetString(0);
                        Assert.That(owned, Is.EqualTo("cached \u03bb"));
                    }
                    else if (row == 10)
                    {
                        Assert.That(reader.IsDBNull(0), Is.True);
                        Assert.That(reader.GetCellError(0), Is.EqualTo(CellError.DIV0));
                    }
                    else
                    {
                        Assert.That(reader.GetDouble(0), Is.EqualTo(row + 0.25));
                    }
                }

                Assert.That(reader.Read(), Is.False);
                Assert.That(owned, Is.EqualTo("cached \u03bb"));
                reader.Reset();
            }
        }

        Assert.That(stream.CanRead, Is.True);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void XmlRowHeightUsesHiddenHeightAndWorksheetDefaultNotCustomHeight(bool singlePass)
    {
        const string rows = """
            <row r="1" ht="24" customHeight="0"><c><v>1</v></c></row>
            <row r="2" hidden="1" ht="50"><c><v>2</v></c></row>
            <row r="3" ht="-7" customHeight="invalid"><c><v>3</v></c></row>
            <row r="5" customHeight="1"><c><v>5</v></c></row>
            <row r="6"/>
            """;
        using var reader = ExcelReaderFactory.CreateReader(
            new MemoryStream(XmlWorkbook(rows, properties: "<sheetFormatPr defaultRowHeight=\"19\"/>")),
            new ExcelReaderConfiguration { SinglePassMode = singlePass });
        double[] heights = [24, 0, 7, 19, 19, 19];
        double?[] values = [1, 2, 3, null, 5, null];
        for (int pass = 0; pass < 2; pass++)
        {
            for (int row = 0; row < heights.Length; row++)
            {
                Assert.That(reader.Read(), Is.True);
                Assert.That(reader.RowHeight, Is.EqualTo(heights[row]));
                if (values[row] is { } value)
                    Assert.That(reader.GetDouble(0), Is.EqualTo(value));
                else
                    Assert.That(reader.IsDBNull(0), Is.True);
            }

            Assert.That(reader.Read(), Is.False);
            reader.Reset();
        }
    }

    [TestCase(false, false)]
    [TestCase(false, true)]
    [TestCase(true, false)]
    [TestCase(true, true)]
    public void XmlRowsResumeWithOwnedValuesAndUnchangedGaps(bool singlePass, bool strict)
    {
        string longText = new('\u03BB', 2000);
        string rows = """
            <row r="1" ht="-24">
              <c r="A1"><v>42</v></c>
              <c t="inlineStr"><is><r><t xml:space="preserve"> &amp; </t></r><r><t>_x0041_</t></r></is></c>
              <c t="inlineStr"><is><t>LONG</t></is></c>
              <c t="e"><v>#DIV/0!</v></c><c/>
              <c t="b"><v>1</v></c><c t="d"><v>2026-10-08</v></c>
              <c t="str"><f>ignored</f><v>cached_x0042_</v></c>
            </row>
            <row r="4" hidden="1"><c t="inlineStr"><is><t>later</t></is></c></row>
            <row/><row r="6"><c r="H6"><v>99</v></c></row>
            """;
        byte[] data = XmlWorkbook(rows.Replace("LONG", longText), strict);
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(data), new ExcelReaderConfiguration { SinglePassMode = singlePass });
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.RowHeight, Is.EqualTo(24));
            Assert.That(reader.FieldCount, Is.EqualTo(8));
            Assert.That(reader.GetDouble(0), Is.EqualTo(42));
            Assert.That(reader.GetString(1), Is.EqualTo(" & A"));
            string owned = reader.GetString(2);
            Assert.That(owned, Is.EqualTo(longText));
            Assert.That(reader.GetCellError(3), Is.EqualTo(CellError.DIV0));
            Assert.That(reader.IsDBNull(3), Is.True);
            Assert.That(reader.IsDBNull(4), Is.True);
            Assert.That(reader.GetBoolean(5), Is.True);
            Assert.That(reader.GetDateTime(6), Is.EqualTo(new DateTime(2026, 10, 8)));
            Assert.That(reader.GetString(7), Is.EqualTo("cachedB"));
            for (int gap = 0; gap < 2; gap++)
            {
                Assert.That(reader.Read(), Is.True);
                Assert.That(reader.RowHeight, Is.EqualTo(15));
                Assert.That(reader.IsDBNull(0), Is.True);
            }

            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.RowHeight, Is.Zero);
            Assert.That(reader.GetString(0), Is.EqualTo("later"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.IsDBNull(0), Is.True);
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(7), Is.EqualTo(99));
            Assert.That(reader.Read(), Is.False);
            Assert.That(owned, Is.EqualTo(longText));
            reader.Reset();
        }
    }

    [TestCase(false)]
    [TestCase(true)]
    public void XmlOutOfOrderHeadersRequireIndexedReading(bool singlePass)
    {
        using var reader = ExcelReaderFactory.CreateReader(
            new MemoryStream(XmlWorkbook("<row r=\"3\"><c><v>3</v></c></row><row r=\"1\"><c><v>1</v></c></row>")),
            new ExcelReaderConfiguration { SinglePassMode = singlePass });
        if (!singlePass)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(0), Is.EqualTo(1));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.IsDBNull(0), Is.True);
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(0), Is.EqualTo(3));
            Assert.That(reader.Read(), Is.False);
            return;
        }

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.RowHeight, Is.Zero);
        Assert.That(reader.FieldCount, Is.EqualTo(singlePass ? 0 : 1));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.RowHeight, Is.EqualTo(15));
        Assert.That(reader.FieldCount, Is.EqualTo(singlePass ? 0 : 1));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(3));
        Assert.That(Assert.Throws<InvalidOperationException>(() => reader.Read())!.Message, Does.Contain("SinglePassMode = false"));
    }

    [TestCase("xlsx", false)]
    [TestCase("xlsb", false)]
    [TestCase("xlsx", true)]
    [TestCase("xlsb", true)]
    public void OpenXmlIndexedReadingMergesWorksheetWideOverlaps(string format, bool distantRevisit)
    {
        int[] order = distantRevisit
            ? [.. Enumerable.Range(0, 70), 0, 1, 75]
            : [0, 3, 1, 4, 2, 5, 7, 6, 8];
        var rows = order.Select((index, ordinal) => new OrderingRow(index, ordinal % 2, ordinal + 0.25, 20 + ordinal, ordinal % 7 == 0)).ToArray();
        rows = [.. rows, new OrderingRow(order.Max() + 3, 0, null, 42, false)];
        using var stream = new MemoryStream(OrderingWorkbook(format, rows));
        using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration { LeaveOpen = true }))
        {
            Assert.That(reader.RowCount, Is.EqualTo(rows.Max(row => row.Index) + 1));
            Assert.That(reader.FieldCount, Is.EqualTo(2));
            Assert.That(reader.Read(), Is.True);
            reader.Reset();
            for (int pass = 0; pass < 2; pass++)
            {
                for (int index = 0; index < reader.RowCount; index++)
                {
                    Assert.That(reader.Read(), Is.True, $"row {index}");
                    var fragments = rows.Where(row => row.Index == index).ToArray();
                    double height = fragments.Length == 0 ? 15 : fragments.Last().Hidden ? 0 : fragments.Last().Height;
                    Assert.That(reader.RowHeight, Is.EqualTo(height), $"row {index}");
                    for (int column = 0; column < 2; column++)
                    {
                        var writes = fragments.Where(row => row.Column == column && row.Value.HasValue).ToArray();
                        Assert.That(reader.GetValue(column), Is.EqualTo(writes.Length == 0 ? DBNull.Value : (object)writes.Last().Value!.Value), $"row {index}, column {column}");
                    }
                }

                Assert.That(reader.Read(), Is.False);
                Assert.That(reader.Read(), Is.False);
                reader.Reset();
            }
        }

        Assert.That(stream.CanRead, Is.True);
    }

    [TestCase("xlsx")]
    [TestCase("xlsb")]
    public void OpenXmlSinglePassMergesCurrentFragmentsButRejectsCompletedRowRevisits(string format)
    {
        OrderingRow[] rows =
        [
            new(0, 0, 1, 24, false),
            new(0, 1, 2, 9, false),
            new(40, 0, 3, 15, false),
            new(0, 0, 4, 15, false),
        ];
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(OrderingWorkbook(format, rows)), new ExcelReaderConfiguration { SinglePassMode = true });
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(0), Is.EqualTo(1));
            Assert.That(reader.GetDouble(1), Is.EqualTo(2));
            Assert.That(reader.RowHeight, Is.EqualTo(9));
            for (int index = 1; index <= 40; index++)
                Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(0), Is.EqualTo(3));
            Assert.That(Assert.Throws<InvalidOperationException>(() => reader.Read())!.Message, Does.Contain("SinglePassMode = false"));
            reader.Reset();
        }
    }

    [TestCase("xlsx")]
    [TestCase("xlsb")]
    public void OpenXmlReorderedLeadingGapsRetainHeights(string format)
    {
        OrderingRow[] rows = [new(4, 0, 4, 24, false), new(2, 0, 2, 18, false)];
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(OrderingWorkbook(format, rows)));
        double[] heights = [0, 15, 18, 15, 24];
        for (int index = 0; index < heights.Length; index++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.RowHeight, Is.EqualTo(heights[index]));
            if (index is 2 or 4)
                Assert.That(reader.GetDouble(0), Is.EqualTo(index));
            else
                Assert.That(reader.IsDBNull(0), Is.True);
        }

        Assert.That(reader.Read(), Is.False);
    }

    [Test]
    public void XmlOrderingScanResolvesImplicitIndices()
    {
        const string rows = """
            <row r="3"><c><v>3</v></c></row>
            <row><c><v>4</v></c></row>
            <row r="1"><c><v>1</v></c></row>
            <row><c><v>2</v></c></row>
            """;
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(XmlWorkbook(rows)));
        for (int index = 1; index <= 4; index++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(0), Is.EqualTo(index));
        }

        Assert.That(reader.Read(), Is.False);
    }

    [Test]
    public void ReorderedBinaryRowsRetainOwnedValuesAndErrors()
    {
        string text = new('\u03bb', 2000);
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(BinaryWorkbook(text, "later", reverseRows: true)));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(4), Is.EqualTo("later"));
        Assert.That(reader.RowHeight, Is.Zero);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.IsDBNull(4), Is.True);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(1), Is.EqualTo(42));
        Assert.That(reader.GetBoolean(2), Is.True);
        Assert.That(reader.GetDouble(3), Is.EqualTo(1.25));
        string owned = reader.GetString(4);
        Assert.That(owned, Is.EqualTo(text));
        Assert.That(reader.GetString(5), Is.EqualTo("shared"));
        Assert.That(reader.GetCellError(6), Is.EqualTo(CellError.DIV0));
        Assert.That(reader.IsDBNull(6), Is.True);
        Assert.That(reader.GetString(7), Is.EqualTo("cached"));
        Assert.That(reader.GetString(8), Is.EqualTo("rich"));
        Assert.That(reader.Read(), Is.False);
        reader.Reset();
        Assert.That(owned, Is.EqualTo(text));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void NextResultCanAbandonAnActiveXmlRowDriver(bool singlePass)
    {
        using var buffer = new MemoryStream();
        byte[] data = XmlWorkbook("<row><c><v>1</v></c></row><row><c><v>2</v></c></row>");
        buffer.Write(data, 0, data.Length);
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            const string workbook = """
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="First" sheetId="1" r:id="sheet1"/><sheet name="Second" sheetId="2" r:id="sheet2"/></sheets>
                </workbook>
                """;
            const string relationships = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="sheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="sheet2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                </Relationships>
                """;
            archive.GetEntry("xl/workbook.xml")!.Delete();
            archive.GetEntry("xl/_rels/workbook.xml.rels")!.Delete();
            WriteEntry(archive, "xl/workbook.xml", workbook);
            WriteEntry(archive, "xl/_rels/workbook.xml.rels", relationships);
        }

        buffer.Position = 0;
        using var reader = ExcelReaderFactory.CreateReader(buffer, new ExcelReaderConfiguration { SinglePassMode = singlePass, LeaveOpen = true });
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(1));
        Assert.That(reader.NextResult(), Is.True);
        Assert.That(reader.Name, Is.EqualTo("Second"));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(1));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(2));
        reader.Reset();
        Assert.That(reader.Name, Is.EqualTo("First"));
        Assert.That(reader.Read(), Is.True);
    }

    [Test]
    public void SinglePassStopsBeforeMalformedNextRowAndResetReopens()
    {
        byte[] data = XmlWorkbook("<row r=\"1\"><c><v>1</v></c></row><row r=\"2\"><c><v>broken</c></row>");
        using var stream = new MemoryStream(data);
        using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration { SinglePassMode = true, LeaveOpen = true }))
        {
            for (int pass = 0; pass < 2; pass++)
            {
                Assert.That(reader.Read(), Is.True);
                Assert.That(reader.GetDouble(0), Is.EqualTo(1));
                reader.Reset();
            }

            Assert.That(reader.Read(), Is.True);
            Assert.Throws<XmlException>(() => reader.Read());
        }

        Assert.That(stream.CanRead, Is.True);
        stream.Position = 0;
        Assert.Throws<XmlException>(() => ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration { LeaveOpen = true }));
        Assert.That(stream.CanRead, Is.True);
    }

    [TestCase("<row r=\"1\"><c><v>1</v></c></row>")]
    [TestCase("<row r=\"1\"/>")]
    public void MalformedFollowingHeaderRetainsLogicalRowLookAhead(string firstRow)
    {
        using var reader = ExcelReaderFactory.CreateReader(
            new MemoryStream(XmlWorkbook(firstRow + "<row r=\"2\" invalid=\"&bad;\"></row>")),
            new ExcelReaderConfiguration { SinglePassMode = true });
        Assert.Throws<XmlException>(() => reader.Read());
        reader.Reset();
        Assert.Throws<XmlException>(() => reader.Read());
    }

    [Test]
    public void LeadingGapsPauseBeforeDecodingThePendingXmlRow()
    {
        using var reader = ExcelReaderFactory.CreateReader(
            new MemoryStream(XmlWorkbook("<row r=\"3\"><c><v>broken</c></row>")),
            new ExcelReaderConfiguration { SinglePassMode = true });
        for (int gap = 0; gap < 2; gap++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.RowHeight, Is.EqualTo(gap == 0 ? 0 : 15));
            Assert.That(reader.FieldCount, Is.Zero);
        }

        Assert.Throws<XmlException>(() => reader.Read());
    }

    [TestCase(false)]
    [TestCase(true)]
    public void ReusedCurrentRowClearsDenseValuesBeforeSparseAndMissingRows(bool singlePass)
    {
        using var reader = ExcelReaderFactory.CreateReader(
            new MemoryStream(XmlWorkbook("""
                <row><c t="inlineStr"><is><t>owned</t></is></c><c><v>2</v></c><c><v>3</v></c></row>
                <row r="3"><c r="C3"><v>4</v></c></row><row/>
                """)),
            new ExcelReaderConfiguration { SinglePassMode = singlePass });
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Read(), Is.True);
            string owned = reader.GetString(0);
            var values = new object[3];
            Assert.That(reader.GetValues(values), Is.EqualTo(3));
            Assert.That(reader.Read(), Is.True);
            Assert.That(Enumerable.Range(0, 3).All(reader.IsDBNull), Is.True);
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.IsDBNull(0), Is.True);
            Assert.That(reader.IsDBNull(1), Is.True);
            Assert.That(reader.GetDouble(2), Is.EqualTo(4));
            Assert.That(reader.Read(), Is.True);
            Assert.That(Enumerable.Range(0, 3).All(reader.IsDBNull), Is.True);
            Assert.That(reader.Read(), Is.False);
            Assert.That(reader.Read(), Is.False);
            Assert.That(owned, Is.EqualTo("owned"));
            Assert.That(values, Is.EqualTo(new object[] { "owned", 2D, 3D }));
            reader.Reset();
        }
    }

    [TestCase(false)]
    [TestCase(true)]
    public void EmptyXmlSheetCompletesWithoutPhantomRows(bool singlePass)
    {
        using var reader = ExcelReaderFactory.CreateReader(
            new MemoryStream(XmlWorkbook(string.Empty)),
            new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Read(), Is.False);
        Assert.That(reader.Read(), Is.False);
        reader.Reset();
        Assert.That(reader.Read(), Is.False);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void BinaryCallbacksOwnLongStringsAndPreserveTypedCells(bool singlePass)
    {
        string first = new('\u03BB', 2000);
        string second = new('\u03A9', 2000);
        using var reader = ExcelReaderFactory.CreateReader(
            new MemoryStream(BinaryWorkbook(first, second)),
            new ExcelReaderConfiguration { SinglePassMode = singlePass });
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.RowHeight, Is.EqualTo(24));
            Assert.That(reader.FieldCount, Is.EqualTo(9));
            Assert.That(reader.IsDBNull(0), Is.True);
            Assert.That(reader.GetDouble(1), Is.EqualTo(42));
            Assert.That(reader.GetBoolean(2), Is.True);
            Assert.That(reader.GetDouble(3), Is.EqualTo(1.25));
            string owned = reader.GetString(4);
            Assert.That(owned, Is.EqualTo(first));
            Assert.That(reader.GetString(5), Is.EqualTo("shared"));
            Assert.That(reader.IsDBNull(6), Is.True);
            Assert.That(reader.GetCellError(6), Is.EqualTo(CellError.DIV0));
            Assert.That(reader.GetString(7), Is.EqualTo("cached"));
            Assert.That(reader.GetString(8), Is.EqualTo("rich"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.IsDBNull(0), Is.True);
            Assert.That(reader.RowHeight, Is.EqualTo(15));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.RowHeight, Is.Zero);
            Assert.That(reader.GetString(4), Is.EqualTo(second));
            Assert.That(reader.Read(), Is.False);
            Assert.That(owned, Is.EqualTo(first));
            reader.Reset();
        }
    }

    [Test]
    public void BinaryMalformedLongStringCanBeResetAfterDecodeFailure()
    {
        using var reader = ExcelReaderFactory.CreateReader(
            new MemoryStream(BinaryWorkbook(new string('a', 2000), new string('b', 2000), malformedSecond: true)),
            new ExcelReaderConfiguration { SinglePassMode = true });
        Assert.That(reader.Read(), Is.True);
        string owned = reader.GetString(4);
        Assert.That(reader.Read(), Is.True);
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.Read());
        reader.Reset();
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetString(4), Is.EqualTo(owned));
    }

    [Test]
    public void StrayMalformedBinaryCellsAreNotSilentlySkipped()
    {
        using var stream = new MemoryStream(BinaryWorkbook(
            new string('a', 2000), new string('b', 2000), malformedSecond: true, omitSheetDataBegin: true));
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelReaderFactory.CreateReader(
            stream, new ExcelReaderConfiguration { SinglePassMode = true, LeaveOpen = true }));
        Assert.That(stream.CanRead, Is.True);
    }

    [TestCase("10x10000.xlsx", false)]
    [TestCase("10x10000.xlsx", true)]
    [TestCase("10x10000.xlsb", false)]
    [TestCase("10x10000.xlsb", true)]
    [TestCase("10x10000.xls", false)]
    [TestCase("10x10000.xls", true)]
    public void PartialReadResetAndNextResultDisposeActiveParser(string fixture, bool singlePass)
    {
        using var stream = Configuration.GetTestWorkbook(fixture);
        using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration { SinglePassMode = singlePass, LeaveOpen = true }))
        {
            Assert.That(reader.Read(), Is.True);
            object first = reader.GetValue(0);
            reader.Reset();
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetValue(0), Is.EqualTo(first));
            _ = reader.NextResult();
            reader.Reset();
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetValue(0), Is.EqualTo(first));
        }

        Assert.That(stream.CanRead, Is.True);
    }

    private static List<string> ReadCallbacks(IWorksheetReader reader)
    {
        using (reader)
        {
            List<string> events = [];
            var handler = new CallbackRecorder(events);
            int pauses = 0;
            WorksheetReadStatus status;
            do
            {
                status = reader.Consume(ref handler);
                if (status == WorksheetReadStatus.Paused)
                    pauses++;
            }
            while (status == WorksheetReadStatus.Paused);

            Assert.That(status, Is.EqualTo(WorksheetReadStatus.EndOfInput));
            Assert.That(reader.Consume(ref handler), Is.EqualTo(WorksheetReadStatus.EndOfInput));
            Assert.That(pauses, Is.EqualTo(events.Count));
            return events;
        }
    }

    private static IWorksheetReader CreateXmlCallbackReader(bool preparing)
    {
        const string xml = """
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetPr codeName="Code"/>
              <sheetFormatPr defaultRowHeight="18"/>
              <dimension ref="A1:B3"/>
              <cols><col min="1" max="2" width="12" customWidth="1"/></cols>
              <mergeCells><mergeCell ref="A1:B1"/></mergeCells>
              <headerFooter><oddHeader>Header</oddHeader></headerFooter>
              <sheetData>
                <row r="1"><c><v>11</v></c><c><v>12</v></c></row>
                <row r="3"><c r="B3"><v>13</v></c></row>
              </sheetData>
            </worksheet>
            """;
        return new XmlWorksheetReader(XmlReader.Create(new StringReader(xml)), preparing ? WorksheetReadMode.Scan : WorksheetReadMode.Materialize, new CommonWorkbook());
    }

    private static IWorksheetReader CreateXlsbCallbackReader(bool preparing)
    {
        using var buffer = new MemoryStream();
        using (var writer = new BinaryWriter(buffer, Encoding.UTF8, leaveOpen: true))
        {
            Record(writer, 0x94, data =>
            {
                data.Write(0);
                data.Write(3);
                data.Write(0);
                data.Write(2);
            });
            Record(writer, 0x3C, data =>
            {
                data.Write(0);
                data.Write(1);
                data.Write(12 * 256);
                data.Write(0);
                data.Write((byte)0);
            });
            Record(writer, 176, data =>
            {
                data.Write(0);
                data.Write(0);
                data.Write(0);
                data.Write(1);
            });
            Record(writer, 0x91, _ => { });
            Record(writer, 0, data => RowHeader(data, 0, 0));
            Cell(writer, 2, 0, data => data.Write(11.0));
            Cell(writer, 2, 1, data => data.Write(12.0));
            Record(writer, 0, data => RowHeader(data, 2, 0));
            Cell(writer, 2, 1, data => data.Write(13.0));
            Record(writer, 0x92, _ => { });
        }

        return new BiffWorksheetReader(new MemoryStream(buffer.ToArray()), preparing ? WorksheetReadMode.Scan : WorksheetReadMode.Materialize, new CommonWorkbook());
    }

    private static byte[] OrderingWorkbook(string format, OrderingRow[] rows)
    {
        if (format == "xlsx")
        {
            var xml = new StringBuilder();
            foreach (var row in rows)
            {
                xml.Append(FormattableString.Invariant($"<row r=\"{row.Index + 1}\" ht=\"{row.Height}\" hidden=\"{(row.Hidden ? 1 : 0)}\">"));
                if (row.Value is { } value)
                    xml.Append(FormattableString.Invariant($"<c r=\"{(char)('A' + row.Column)}{row.Index + 1}\"><v>{value}</v></c>"));
                xml.Append("</row>");
            }

            return XmlWorkbook(xml.ToString());
        }

        using var buffer = new MemoryStream();
        byte[] data = BinaryWorkbook("first", "second");
        buffer.Write(data, 0, data.Length);
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            archive.GetEntry("xl/worksheets/sheet1.bin")!.Delete();
            using var writer = new BinaryWriter(archive.CreateEntry("xl/worksheets/sheet1.bin").Open());
            Record(writer, 0x91, _ => { });
            foreach (var row in rows)
            {
                Record(writer, 0, header =>
                {
                    header.Write(row.Index);
                    header.Write(0);
                    header.Write((ushort)(row.Height * 20));
                    header.Write((byte)0);
                    header.Write((byte)(0b100000 | (row.Hidden ? 0b10000 : 0)));
                });
                if (row.Value is { } value)
                    Cell(writer, 5, row.Column, cell => cell.Write(value));
            }

            Record(writer, 0x92, _ => { });
        }

        return buffer.ToArray();
    }

    private static byte[] XmlWorkbook(string rows, bool strict = false, string properties = "")
    {
        using var buffer = new MemoryStream();
        byte[] data = AllocationTestWorkbook.CreateXlsx([], false);
        buffer.Write(data, 0, data.Length);
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            archive.GetEntry("xl/worksheets/sheet1.xml")!.Delete();
            string ns = strict ? "http://purl.oclc.org/ooxml/spreadsheetml/main" : "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            WriteEntry(archive, "xl/worksheets/sheet1.xml", "<worksheet xmlns=\"" + ns + "\">" + properties + "<sheetData>" + rows + "</sheetData></worksheet>");
        }

        return buffer.ToArray();
    }

    private static byte[] XlsWorkbook(bool overlapping)
    {
        using var buffer = new MemoryStream();
        using var writer = new BinaryWriter(buffer, Encoding.Unicode, leaveOpen: true);
        XlsRecord(writer, 0x809, data =>
        {
            data.Write((ushort)0x600);
            data.Write((ushort)0x10);
            data.Write(new byte[12]);
        });
        for (int row = 0; row < 40; row++)
        {
            XlsRecord(writer, 0x208, data =>
            {
                data.Write((ushort)row);
                data.Write((ushort)0);
                data.Write((ushort)2);
                data.Write((ushort)(row == 1 ? 480 : 255));
                data.Write(0);
                data.Write(0);
            });
        }

        int[] order = overlapping ? [0, 33, .. Enumerable.Range(1, 32), .. Enumerable.Range(34, 6)] : [.. Enumerable.Range(0, 40)];
        foreach (int row in order)
        {
            if (row == 1)
            {
                XlsRecord(writer, 0xbd, data =>
                {
                    data.Write((ushort)row);
                    data.Write((ushort)0);
                    for (int number = 7; number <= 8; number++)
                    {
                        data.Write((ushort)0);
                        data.Write((number << 2) | 2);
                    }

                    data.Write((ushort)1);
                });
            }
            else
            {
                XlsRecord(writer, row == 2 ? (ushort)6 : row == 10 ? (ushort)0x205 : (ushort)0x203, data =>
                {
                    data.Write((ushort)row);
                    data.Write((ushort)0);
                    data.Write((ushort)0);
                    if (row == 2)
                        data.Write(new byte[] { 0, 0, 0, 0, 0, 0, 255, 255 });
                    else if (row == 10)
                        data.Write(new byte[] { (byte)CellError.DIV0, 1 });
                    else
                        data.Write(row + 0.25);
                });
                if (row == 2)
                {
                    XlsRecord(writer, 0x207, data =>
                    {
                        const string value = "cached \u03bb";
                        data.Write((ushort)value.Length);
                        data.Write((byte)1);
                        data.Write(Encoding.Unicode.GetBytes(value));
                    });
                }
            }
        }

        XlsRecord(writer, 0xa, _ => { });
        return buffer.ToArray();
    }

    private static void XlsRecord(BinaryWriter writer, ushort id, Action<BinaryWriter> value)
    {
        using var buffer = new MemoryStream();
        using (var data = new BinaryWriter(buffer, Encoding.Unicode, leaveOpen: true))
            value(data);
        writer.Write(id);
        writer.Write((ushort)buffer.Length);
        writer.Write(buffer.ToArray());
    }

    private static byte[] BinaryWorkbook(string first, string second, bool malformedSecond = false, bool omitSheetDataBegin = false, bool reverseRows = false)
    {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            const string relationships = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="sheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/>
                  <Relationship Id="sst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>
                </Relationships>
                """;
            WriteEntry(archive, "xl/_rels/workbook.bin.rels", relationships);
            using (var writer = new BinaryWriter(archive.CreateEntry("xl/workbook.bin").Open()))
            {
                Record(writer, 0x83, _ => { });
                Record(writer, 0x9C, data =>
                {
                    data.Write(0);
                    data.Write(1);
                    WideString(data, "sheet1");
                    WideString(data, "Sheet1");
                });
                Record(writer, 0x84, _ => { });
            }

            using (var writer = new BinaryWriter(archive.CreateEntry("xl/sharedStrings.bin").Open()))
            {
                Record(writer, 0x9F, data =>
                {
                    data.Write(1);
                    data.Write(1);
                });
                Record(writer, 0x13, data =>
                {
                    data.Write((byte)0);
                    WideString(data, "shared");
                });
                Record(writer, 0xA0, _ => { });
            }

            using (var writer = new BinaryWriter(archive.CreateEntry("xl/worksheets/sheet1.bin").Open()))
            {
                if (!omitSheetDataBegin)
                    Record(writer, 0x91, _ => { });
                Record(writer, 0, data => RowHeader(data, reverseRows ? 2 : 0, 0b100000));
                Cell(writer, 1, 0, _ => { });
                Cell(writer, 2, 1, data => data.Write((42 << 2) | 2));
                Cell(writer, 4, 2, data => data.Write((byte)1));
                Cell(writer, 5, 3, data => data.Write(1.25));
                Cell(writer, 6, 4, data => WideString(data, first));
                Cell(writer, 7, 5, data => data.Write(0));
                Cell(writer, 0x0B, 6, data => data.Write((byte)CellError.DIV0));
                Cell(writer, 8, 7, data => WideString(data, "cached"));
                Cell(writer, 0x3E, 8, data =>
                {
                    data.Write((byte)0);
                    WideString(data, "rich");
                });
                Record(writer, 0, data => RowHeader(data, reverseRows ? 0 : 2, 0b110000));
                Cell(writer, 6, 4, data =>
                {
                    data.Write(malformedSecond ? int.MaxValue : second.Length);
                    data.Write(Encoding.Unicode.GetBytes(second));
                });
                Record(writer, 0x92, _ => { });
            }
        }

        return buffer.ToArray();
    }

    private static void RowHeader(BinaryWriter writer, int index, byte flags)
    {
        writer.Write(index);
        writer.Write(0);
        writer.Write((ushort)480);
        writer.Write((byte)0);
        writer.Write(flags);
    }

    private static void WideString(BinaryWriter writer, string value)
    {
        writer.Write(value.Length);
        writer.Write(Encoding.Unicode.GetBytes(value));
    }

    private static void Cell(BinaryWriter writer, uint id, int column, Action<BinaryWriter> value) =>
        Record(writer, id, data =>
        {
            data.Write(column);
            data.Write(0);
            value(data);
        });

    private static void Record(BinaryWriter writer, uint id, Action<BinaryWriter> value)
    {
        using var buffer = new MemoryStream();
        using (var data = new BinaryWriter(buffer, Encoding.UTF8, leaveOpen: true))
            value(data);
        Variable(writer, id);
        Variable(writer, (uint)buffer.Length);
        writer.Write(buffer.ToArray());
    }

    private static void Variable(BinaryWriter writer, uint value)
    {
        while (value >= 128)
        {
            writer.Write((byte)((value & 127) | 128));
            value >>= 7;
        }

        writer.Write((byte)value);
    }

    private static void WriteEntry(ZipArchive archive, string name, string text)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
        writer.Write(text);
    }

    private readonly struct OrderingRow(int index, int column, double? value, double height, bool hidden)
    {
        public int Index { get; } = index;

        public int Column { get; } = column;

        public double? Value { get; } = value;

        public double Height { get; } = height;

        public bool Hidden { get; } = hidden;
    }

    private struct CallbackRecorder(List<string> events) : IWorksheetHandler
    {
        private int _rowIndex;

        public WorksheetReadControl OnDimension(CellRange range) => Add("dimension");

        public WorksheetReadControl OnCodeName(string codeName) => Add("code");

        public WorksheetReadControl OnDefaultRowHeight(double? height) => Add("height");

        public WorksheetReadControl OnColumn(Column column) => Add("column");

        public WorksheetReadControl OnMergeCell(CellRange range) => Add("merge");

        public WorksheetReadControl OnHeaderFooter(HeaderFooter headerFooter) => Add("header");

        public WorksheetReadControl OnSheetDataBegin() => Add("begin");

        public WorksheetReadControl OnScanResult(int maxRowIndex, int maxColumnIndex) => Add("scan");

        public WorksheetReadControl OnRow(int index, bool hidden, double? height)
        {
            _rowIndex = index;
            return Add($"row:{index}");
        }

        public WorksheetReadControl OnCell(in ExcelDataReader.Core.Cell cell) =>
            Add($"cell:{_rowIndex}:{cell.ColumnIndex}");

        public WorksheetReadControl OnRowEnd() => WorksheetReadControl.Continue;

        private WorksheetReadControl Add(string value)
        {
            events.Add(value);
            return WorksheetReadControl.Pause;
        }
    }

    private struct ThrowingCallbackHandler : IWorksheetHandler
    {
        public WorksheetReadControl OnDimension(CellRange range) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnCodeName(string codeName) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnDefaultRowHeight(double? height) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnColumn(Column column) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnMergeCell(CellRange range) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnHeaderFooter(HeaderFooter headerFooter) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnSheetDataBegin() => WorksheetReadControl.Continue;

        public WorksheetReadControl OnScanResult(int maxRowIndex, int maxColumnIndex) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnRow(int index, bool hidden, double? height) => WorksheetReadControl.Continue;

        public WorksheetReadControl OnCell(in ExcelDataReader.Core.Cell cell) =>
            throw new InvalidOperationException("Test callback failure.");

        public WorksheetReadControl OnRowEnd() => WorksheetReadControl.Continue;
    }
}