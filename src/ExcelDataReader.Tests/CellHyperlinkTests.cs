#nullable enable

using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelDataReader.Core;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Tests;

public sealed class CellHyperlinkTests
{
    private const int Rows = 700;
    private const string Address = @"..\Reports\file%23name.xlsx#'Sheet #2'!A1";
    private const string Tooltip = "Tip & <text> \u03bb";
    private static readonly string[] Formats = ["xls", "xlsx", "xlsb", "xml"];
    private static readonly string[] ExcelSavedFiles =
    [
        "Hyperlinks.xls", "Hyperlinks.xlsx", "Hyperlinks.xlsb", "Hyperlinks.xml",
        "HyperlinksPassword.xls", "HyperlinksPassword.xlsx", "HyperlinksPassword.xlsb",
    ];

    [Test]
    public void ExcelSavedHyperlinksPreserveDestinationsAndRows(
        [ValueSource(nameof(ExcelSavedFiles))] string file,
        [Values(false, true)] bool singlePass)
    {
        bool xml = file.EndsWith(".xml", StringComparison.Ordinal);
        string directory = xml ? @"C:\Reports\" : "../../../../../../Reports/";
        var expected = new (string? Value, string? Address, string? Tooltip)[]
        {
            ("Website fragment", "https://example.com/report#section", "Website tooltip"),
            ("External workbook cell", directory + "Budget.xlsx#'Summary'!A1", "File tooltip"),
            ("Internal cell", "#'Summary'!A1", "Internal tooltip"),
            ("Internal name", "#SalesTotals", "Name tooltip"),
            ("Email", "mailto:example@example.com", "Email tooltip"),
            ("External workbook name", directory + "Budget.xlsx#SalesTotals", "External name tooltip"),
            ("Tooltip only", string.Empty, "Only a tooltip"),
            ("Escaped address", "https://example.com/a%23b#one#two", null),
            ("Unicode path", directory + "\u5831\u544a.xlsx", "Unicode tooltip"),
            ("Plain URL https://example.com/", null, null),

            // Excel's XML export also stores an explicit HRef for the HYPERLINK formula.
            ("Formula label", xml ? "https://example.com/formula" : null, null),
            ("Range label", "https://example.com/range", "Range tooltip"),
            ("Range label", "https://example.com/range", "Range tooltip"),
            ("Range label", "https://example.com/range", "Range tooltip"),
            ("Merged label", "https://example.com/merged", "Merged tooltip"),
        };

        using var source = Configuration.GetTestWorkbook(file);
        using var reader = ExcelReaderFactory.CreateReader(source, new ExcelReaderConfiguration
        {
            Password = "password",
            SinglePassMode = singlePass,
        });
        CellHyperlink? retained = null;
        CellHyperlink? range = null;
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Name, Is.EqualTo("Links"));
            Assert.Throws<InvalidOperationException>(() => reader.GetHyperlink(0));
            for (int row = 0; row < 700; row++)
            {
                Assert.That(reader.Read(), Is.True, $"row {row}");
                Assert.That(reader.GetDouble(2), Is.EqualTo(row + 1));
                Assert.That(reader.IsDBNull(1), Is.True);
                long position = source.Position;
                var link = reader.GetHyperlink(0);
                Assert.That(reader.GetHyperlink(1), Is.Null);
                Assert.That(reader.GetHyperlink(2), Is.Null);
                if (row < expected.Length)
                {
                    Assert.That(reader.GetValue(0), Is.EqualTo(expected[row].Value));
                    if (expected[row].Address == null)
                        Assert.That(link, Is.Null);
                    else
                        AssertLink(link, expected[row].Address!, expected[row].Tooltip);
                }
                else
                {
                    Assert.That(reader.IsDBNull(0), Is.True);
                    Assert.That(link, Is.Null);
                }

                if (row == 0)
                {
                    if (pass == 0)
                        retained = link;
                    Assert.That(link, Is.SameAs(retained));
                }

                if (row == 11)
                    range = link;
                if (!xml && row is 12 or 13)
                    Assert.That(link, Is.SameAs(range));
                if (file.EndsWith(".xls", StringComparison.Ordinal) || xml)
                    Assert.That(source.Position, Is.EqualTo(position));
            }

            Assert.That(reader.FieldCount, Is.EqualTo(3));
            if (!singlePass)
                Assert.That(reader.RowCount, Is.EqualTo(700));
            Assert.That(reader.Read(), Is.False);
            Assert.Throws<InvalidOperationException>(() => reader.GetHyperlink(0));
            Assert.That(reader.NextResult(), Is.True);
            Assert.That(reader.Name, Is.EqualTo("Summary"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo("Target"));
            Assert.That(reader.GetHyperlink(0), Is.Null);
            Assert.That(reader.Read(), Is.False);
            Assert.That(reader.NextResult(), Is.False);
            reader.Reset();
        }

        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables.Count, Is.EqualTo(2));
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(700));
        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(3));
        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo("Website fragment"));
        Assert.That(dataSet.Tables[0].Rows[699][2], Is.EqualTo(700));
        Assert.That(dataSet.Tables[1].Rows[0][0], Is.EqualTo("Target"));
        reader.Close();
        AssertLink(retained, "https://example.com/report#section", "Website tooltip");
    }

    [Test]
    public void LazyHyperlinksPreserveValuesAndLifecycle(
        [ValueSource(nameof(Formats))] string format,
        [Values(false, true)] bool singlePass)
    {
        byte[] bytes = CreateWorkbook(format);
        using var source = new MemoryStream(bytes);
        using var reader = ExcelReaderFactory.CreateReader(source, new ExcelReaderConfiguration { SinglePassMode = singlePass, LeaveOpen = true });
        using var baseline = ExcelReaderFactory.CreateReader(new MemoryStream(bytes), new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.Throws<InvalidOperationException>(() => reader.GetHyperlink(0));
        CellHyperlink? retained = null;
        for (int row = 0; row < Rows; row++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(baseline.Read(), Is.True);
            Assert.That(reader.GetValue(0), Is.EqualTo(baseline.GetValue(0)));
            Assert.That(reader.GetValue(1), Is.EqualTo(baseline.GetValue(1)));
            long position = source.Position;
            if (row == 13)
            {
                retained = reader.GetHyperlink(0);
                AssertLink(retained, Address, Tooltip);
                Assert.That(reader.GetHyperlink(0), Is.SameAs(retained));
                var other = reader.GetHyperlink(1);
                AssertLink(other, Address, "Other tooltip");
                Assert.That(other, Is.Not.SameAs(retained));
                Assert.Throws<IndexOutOfRangeException>(() => reader.GetHyperlink(-1));
                Assert.Throws<IndexOutOfRangeException>(() => reader.GetHyperlink(reader.FieldCount));
            }
            else if (row == 14)
            {
                AssertLink(reader.GetHyperlink(0), "#SalesTotals", null);
            }
            else if (row == 15)
            {
                AssertLink(reader.GetHyperlink(0), Address, Tooltip);
                if (format != "xml")
                    Assert.That(reader.GetHyperlink(0), Is.SameAs(retained));
            }
            else if (row == 16)
            {
                AssertLink(reader.GetHyperlink(0), string.Empty, Tooltip);
            }
            else if (row > 16)
            {
                Assert.That(reader.GetHyperlink(0), Is.Null);
            }

            if (format is "xls" or "xml")
                Assert.That(source.Position, Is.EqualTo(position), $"row {row}");
        }

        Assert.That(reader.FieldCount, Is.EqualTo(2));
        if (!singlePass)
            Assert.That(reader.RowCount, Is.EqualTo(Rows));
        Assert.That(reader.Read(), Is.False);
        Assert.Throws<InvalidOperationException>(() => reader.GetHyperlink(0));
        Assert.That(reader.NextResult(), Is.False);
        reader.Reset();
        for (int row = 0; row <= 13; row++)
            Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetHyperlink(0), Is.SameAs(retained));
        reader.Close();
        Assert.That(source.CanRead, Is.True);
        Assert.Throws<InvalidOperationException>(() => reader.GetHyperlink(0));
        AssertLink(retained, Address, Tooltip);
    }

    [TestCase("xlsx")]
    [TestCase("xlsb")]
    [TestCase("xls")]
    public void InvalidMetadataIsDeferredAndFailedLoadsAreNotCached(string format)
    {
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(CreateWorkbook(format, invalidRange: true)));
        Assert.That(reader.Read(), Is.True);
        Assert.Throws<ExcelReaderException>(() => reader.GetHyperlink(0));
        Assert.Throws<ExcelReaderException>(() => reader.GetHyperlink(0));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(1));
    }

    [Test]
    public void TruncatedBinaryHyperlinkStringsFailOnlyWhenRequested()
    {
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(CreateWorkbook("xlsb", invalidString: true)));
        Assert.That(reader.Read(), Is.True);
        Assert.Throws<ExcelReaderException>(() => reader.GetHyperlink(0));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(1));
    }

    [TestCase("Issue242_StdRc4PwdPassword.xls")]
    [TestCase("Issue242_XorPwdPassword.xls")]
    [TestCase("agile_AES128_SHA1_CBC_pwd_password.xlsx")]
    [TestCase("agile_AES128_SHA1_CBC_pwd_password.xlsb")]
    [TestCase("standard_AES128_SHA1_ECB_pwd_password.xlsx")]
    public void LazyScanPreservesEncryptedRows(string file)
    {
        using var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook(file), new ExcelReaderConfiguration { Password = "password" });
        using var baseline = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook(file), new ExcelReaderConfiguration { Password = "password" });
        bool next;
        do
        {
            while (reader.Read())
            {
                Assert.That(baseline.Read(), Is.True);
                if (reader.FieldCount > 0)
                    _ = reader.GetHyperlink(0);
                for (int column = 0; column < reader.FieldCount; column++)
                    Assert.That(reader.GetValue(column), Is.EqualTo(baseline.GetValue(column)));
            }

            Assert.That(baseline.Read(), Is.False);
            next = reader.NextResult();
            Assert.That(baseline.NextResult(), Is.EqualTo(next));
        }
        while (next);
    }

    [Test]
    public void CsvReturnsNoHyperlinkWithoutChangingUrlText()
    {
        using var reader = ExcelReaderFactory.CreateCsvReader(
            new MemoryStream(Encoding.UTF8.GetBytes("https://example.com/#anchor")),
            new ExcelReaderConfiguration { AutodetectSeparators = [","] });
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetHyperlink(0), Is.Null);
        Assert.That(reader.GetString(0), Is.EqualTo("https://example.com/#anchor"));
    }

    [Test]
    public void StrictNamespacesAndRelationshipTypesAreRespected()
    {
        byte[] strict = RewritePackage(CreateWorkbook("xlsx"), (name, text) => text
            .Replace("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main")
            .Replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"));
        using (var reader = ExcelReaderFactory.CreateReader(new MemoryStream(strict)))
        {
            for (int row = 0; row <= 13; row++)
                Assert.That(reader.Read(), Is.True);
            AssertLink(reader.GetHyperlink(0), Address, Tooltip);
        }

        byte[] unrelated = RewritePackage(strict, (name, text) => name.Contains("worksheets/_rels/")
            ? text.Replace("/hyperlink", "/unrelated")
            : text);
        using (var reader = ExcelReaderFactory.CreateReader(new MemoryStream(unrelated)))
        {
            for (int row = 0; row <= 13; row++)
                Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetHyperlink(0), Is.Null);
            Assert.That(reader.Read(), Is.True);
            AssertLink(reader.GetHyperlink(0), "#SalesTotals", null);
        }
    }

    [TestCase("r:id=\"external\"", "r:id=\"missing\"")]
    [TestCase("Target=", "MissingTarget=")]
    public void MissingExternalDestinationIsNotReinterpretedAsInternal(string original, string replacement)
    {
        byte[] bytes = RewritePackage(CreateWorkbook("xlsx"), (name, text) => name.Contains("worksheets/")
            ? text.Replace(original, replacement)
            : text);
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(bytes));
        for (int row = 0; row <= 13; row++)
            Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetHyperlink(0), Is.Null);
    }

    [Test]
    public void SpreadsheetXmlIndicesDoNotPropagateMergedCellLinks()
    {
        const string xml = """
            <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
                      xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
                      xmlns:x="urn:schemas-microsoft-com:office:excel">
              <Worksheet ss:Name="First"><Table>
                <Row ss:Index="3">
                  <Cell ss:MergeAcross="1" ss:HRef="#UnknownName" x:HRefScreenTip="tip"><Data ss:Type="String">label</Data></Cell>
                  <Cell ss:HRef="relative file"><Data ss:Type="Number">42</Data></Cell>
                </Row>
                <Row><Cell><Data ss:Type="Number">43</Data></Cell></Row>
              </Table></Worksheet>
              <Worksheet ss:Name="Second"><Table>
                <Row><Cell ss:HRef="" x:HRefScreenTip="only tip"><Data ss:Type="String">other</Data></Cell></Row>
              </Table></Worksheet>
            </Workbook>
            """;
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        using var reader = ExcelReaderFactory.CreateReader(source);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetHyperlink(0), Is.Null);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.Read(), Is.True);
        AssertLink(reader.GetHyperlink(0), "#UnknownName", "tip");
        Assert.That(reader.GetHyperlink(1), Is.Null);
        AssertLink(reader.GetHyperlink(2), "relative file", null);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(43));
        Assert.That(reader.GetHyperlink(0), Is.Null);
        Assert.That(reader.NextResult(), Is.True);
        Assert.Throws<InvalidOperationException>(() => reader.GetHyperlink(0));
        Assert.That(reader.Read(), Is.True);
        AssertLink(reader.GetHyperlink(0), string.Empty, "only tip");
    }

    [TestCase("xml")]
    [TestCase("xls")]
    public void CallerScanErrorsAreNotHiddenOrCached(string format)
    {
        using var source = new FailingScanStream(CreateWorkbook(format));
        using var reader = ExcelReaderFactory.CreateReader(source, new ExcelReaderConfiguration { LeaveOpen = true });
        for (int row = 0; row <= 13; row++)
            Assert.That(reader.Read(), Is.True);
        long position = source.Position;
        source.Fail = true;
        Assert.That(Assert.Throws<IOException>(() => reader.GetHyperlink(0)), Is.SameAs(source.Error));
        Assert.That(source.Position, Is.EqualTo(position));
        source.Fail = false;
        AssertLink(reader.GetHyperlink(0), Address, Tooltip);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(14));
    }

    [Test]
    public void RangeIndexPreservesOrdinalWithoutExpandingArea()
    {
        var store = new HyperlinkStore();
        var first = new CellHyperlink("first", null);
        var last = new CellHyperlink("last", null);
        store.Add(new CellRange(0, 0, 16383, 1048575), first);
        store.Add(new CellRange(2, 5, 4, 8), last);
        store.Add(new CellRange(10, 2, 10, 3), new CellHyperlink("middle", null));
        Assert.That(store.Find(7, 3), Is.SameAs(last));
        Assert.That(store.Find(1048575, 16383), Is.SameAs(first));
        Assert.That(store.Find(1, 1), Is.SameAs(first));
    }

    [Test]
    public void RangeIndexFindsLastStoredLinkAcrossReorderedLookups()
    {
        var store = new HyperlinkStore();
        var entries = new List<(CellRange Range, CellHyperlink Link)>();
        var random = new Random(17);
        for (int i = 0; i < 100; i++)
        {
            int row = random.Next(30);
            int column = random.Next(10);
            var range = new CellRange(column, row, column + random.Next(5), row + random.Next(10));
            var link = new CellHyperlink(i.ToString(CultureInfo.InvariantCulture), null);
            entries.Add((range, link));
            store.Add(range, link);
        }

        for (int i = 0; i < 300; i++)
        {
            int row = random.Next(40);
            int column = random.Next(15);
            CellHyperlink? expected = null;
            foreach (var entry in entries)
            {
                if (entry.Range.FromRow <= row && entry.Range.ToRow >= row &&
                    entry.Range.FromColumn <= column && entry.Range.ToColumn >= column)
                    expected = entry.Link;
            }

            Assert.That(store.Find(row, column), Is.SameAs(expected), $"row {row}, column {column}");
        }
    }

    [TestCase("A0")]
    [TestCase("XFE1")]
    [TestCase("A1048577")]
    [TestCase("AAAAAAAAAAAA1")]
    [TestCase("A1:B2:C3")]
    [TestCase("B2:A1")]
    public void InvalidRangeCannotBecomeA1(string reference)
    {
        Assert.Throws<ExcelReaderException>(() => HyperlinkStore.ParseRange(reference));
    }

    [TestCase("https://example.com/a%23b", "one#two", "https://example.com/a%23b#one#two")]
    [TestCase("https://example.com/a#first", "second", "https://example.com/a#first#second")]
    [TestCase(null, "SalesTotals", "#SalesTotals")]
    public void JoiningPreservesStoredComponents(string? address, string location, string expected)
    {
        Assert.That(HyperlinkStore.JoinAddress(address, location), Is.EqualTo(expected));
    }

    private static void AssertLink(CellHyperlink? link, string address, string? tooltip)
    {
        Assert.That(link, Is.Not.Null);
        Assert.That(link!.Address, Is.EqualTo(address));
        Assert.That(link.Tooltip, Is.EqualTo(tooltip));
    }

    private static byte[] CreateWorkbook(string format, bool invalidRange = false, bool invalidString = false)
    {
        var links = new[]
        {
            (Range: new CellRange(0, 13, 0, 15), Address, Tooltip: (string?)Tooltip),
            (Range: new CellRange(1, 13, 1, 13), Address, Tooltip: (string?)"Other tooltip"),
            (Range: new CellRange(0, 14, 0, 14), Address: "#SalesTotals", Tooltip: (string?)null),
            (Range: new CellRange(0, 16, 0, 16), Address: string.Empty, Tooltip: (string?)Tooltip),
            (Range: new CellRange(0, Rows + 10, 0, Rows + 10), Address: "outside", Tooltip: (string?)null)
        };
        if (format == "xml")
        {
            var xml = new StringBuilder("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><Worksheet ss:Name=\"Links\"><Table>");
            for (int row = 0; row < Rows; row++)
            {
                xml.Append("<Row>");
                for (int column = 0; column < 2; column++)
                {
                    var matches = links.Where(link => link.Range.FromRow <= row && link.Range.ToRow >= row && link.Range.FromColumn == column).ToArray();
                    xml.Append("<Cell");
                    if (matches.Length > 0)
                    {
                        var link = matches[matches.Length - 1];
                        xml.Append($" ss:HRef=\"{Escape(link.Address)}\"");
                        if (link.Tooltip != null)
                            xml.Append($" x:HRefScreenTip=\"{Escape(link.Tooltip)}\"");
                    }

                    xml.AppendFormat(CultureInfo.InvariantCulture, "><Data ss:Type=\"Number\">{0}</Data></Cell>", row + column);
                }

                xml.Append("</Row>");
            }

            xml.Append("</Table></Worksheet></Workbook>");
            return Encoding.UTF8.GetBytes(xml.ToString());
        }

        using var output = new MemoryStream();
        if (format == "xls")
        {
            using var writer = new BinaryWriter(output, Encoding.Unicode, leaveOpen: true);
            Record(
                writer,
                0x809,
                data =>
            {
                data.Write((ushort)0x600);
                data.Write((ushort)0x10);
                data.Write(new byte[12]);
            },
                binary: false);
            for (int row = 0; row < Rows; row++)
            {
                for (int column = 0; column < 2; column++)
                {
                    Record(
                        writer,
                        0x203,
                        data =>
                    {
                        data.Write((ushort)row);
                        data.Write((ushort)column);
                        data.Write((ushort)0);
                        data.Write((double)(row + column));
                    },
                        binary: false);
                }
            }

            foreach (var link in links)
            {
                Record(
                    writer,
                    0x1B8,
                    data =>
                {
                    WriteRange(data, link.Range, binary: false, invalidRange);
                    data.Write(new Guid("79eac9d0-baf9-11ce-8c82-00aa004ba90b").ToByteArray());
                    data.Write(2);
                    data.Write(0x101);
                    WideString(data, link.Address + "\0");
                },
                    binary: false);
                if (link.Tooltip != null)
                {
                    Record(
                        writer,
                        0x800,
                        data =>
                    {
                        data.Write((ushort)0x800);
                        WriteRange(data, link.Range, binary: false, invalidRange: false);
                        data.Write(Encoding.Unicode.GetBytes(link.Tooltip + "\0"));
                    },
                        binary: false);
                }
            }

            Record(writer, 0xA, _ => { }, binary: false);
            return output.ToArray();
        }

        bool binary = format == "xlsb";
        string extension = binary ? "bin" : "xml";
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(zip, $"xl/_rels/workbook.{extension}.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"sheet\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/data.{extension}\"/></Relationships>");
            WriteEntry(zip, $"xl/worksheets/_rels/data.{extension}.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"external\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{Escape(Address.Substring(0, Address.IndexOf('#')))}\" TargetMode=\"External\"/></Relationships>");
            if (binary)
            {
                using (var writer = new BinaryWriter(zip.CreateEntry("xl/workbook.bin").Open()))
                {
                    Record(writer, 0x83, _ => { });
                    Record(writer, 0x9C, data =>
                    {
                        data.Write(0);
                        data.Write(1);
                        WideString(data, "sheet");
                        WideString(data, "Links");
                    });
                    Record(writer, 0x84, _ => { });
                }

                using (var writer = new BinaryWriter(zip.CreateEntry("xl/worksheets/data.bin").Open()))
                {
                    Record(writer, 0x91, _ => { });
                    for (int row = 0; row < Rows; row++)
                    {
                        Record(writer, 0, data =>
                        {
                            data.Write(row);
                            data.Write(0);
                            data.Write((ushort)300);
                            data.Write((ushort)0);
                        });
                        for (int column = 0; column < 2; column++)
                        {
                            Record(writer, 5, data =>
                            {
                                data.Write(column);
                                data.Write(0);
                                data.Write((double)(row + column));
                            });
                        }
                    }

                    Record(writer, 0x92, _ => { });
                    foreach (var link in links)
                    {
                        Record(writer, 494, data =>
                        {
                            WriteRange(data, link.Range, binary: true, invalidRange);
                            if (invalidString)
                            {
                                data.Write(1000000);
                                return;
                            }

                            if (link.Address == Address)
                                WideString(data, "external");
                            else
                                data.Write(uint.MaxValue);
                            WideString(data, link.Address == Address ? Address.Substring(Address.IndexOf('#') + 1) : link.Address.TrimStart('#'));
                            WideString(data, link.Tooltip ?? string.Empty);
                            WideString(data, "Not the cell value");
                        });
                    }
                }
            }
            else
            {
                WriteEntry(zip, "xl/workbook.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Links\" sheetId=\"1\" r:id=\"sheet\"/></sheets></workbook>");
                var xml = new StringBuilder("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData>");
                for (int row = 0; row < Rows; row++)
                    xml.AppendFormat(CultureInfo.InvariantCulture, "<row r=\"{0}\"><c r=\"A{0}\"><v>{1}</v></c><c r=\"B{0}\"><v>{0}</v></c></row>", row + 1, row);
                xml.Append("</sheetData><hyperlinks>");
                foreach (var link in links)
                {
                    string reference = invalidRange ? "A0" : $"{(char)('A' + link.Range.FromColumn)}{link.Range.FromRow + 1}:{(char)('A' + link.Range.ToColumn)}{link.Range.ToRow + 1}";
                    xml.Append($"<hyperlink ref=\"{reference}\"");
                    if (link.Address == Address)
                        xml.Append(" r:id=\"external\"");
                    xml.Append($" location=\"{Escape(link.Address == Address ? Address.Substring(Address.IndexOf('#') + 1) : link.Address.TrimStart('#'))}\"");
                    if (link.Tooltip != null)
                        xml.Append($" tooltip=\"{Escape(link.Tooltip)}\"");
                    xml.Append(" display=\"Not the cell value\"/>");
                }

                xml.Append("</hyperlinks></worksheet>");
                WriteEntry(zip, "xl/worksheets/data.xml", xml.ToString());
            }
        }

        return output.ToArray();
    }

    private static void WriteRange(BinaryWriter writer, CellRange range, bool binary, bool invalidRange)
    {
        if (binary)
        {
            writer.Write(invalidRange ? -1 : range.FromRow);
            writer.Write(range.ToRow);
            writer.Write(range.FromColumn);
            writer.Write(range.ToColumn);
        }
        else
        {
            writer.Write((ushort)(invalidRange ? range.ToRow + 1 : range.FromRow));
            writer.Write((ushort)range.ToRow);
            writer.Write((ushort)range.FromColumn);
            writer.Write((ushort)range.ToColumn);
        }
    }

    private static void Record(BinaryWriter writer, uint id, Action<BinaryWriter> write, bool binary = true)
    {
        using var payload = new MemoryStream();
        using (var data = new BinaryWriter(payload, Encoding.Unicode, leaveOpen: true))
            write(data);
        if (binary)
        {
            WriteVariable(writer, id);
            WriteVariable(writer, (uint)payload.Length);
        }
        else
        {
            writer.Write((ushort)id);
            writer.Write((ushort)payload.Length);
        }

        writer.Write(payload.ToArray());
    }

    private static void WriteVariable(BinaryWriter writer, uint value)
    {
        while (value >= 128)
        {
            writer.Write((byte)(value | 128));
            value >>= 7;
        }

        writer.Write((byte)value);
    }

    private static void WideString(BinaryWriter writer, string text)
    {
        writer.Write(text.Length);
        writer.Write(Encoding.Unicode.GetBytes(text));
    }

    private static void WriteEntry(ZipArchive zip, string name, string text)
    {
        using var writer = new StreamWriter(zip.CreateEntry(name).Open(), Encoding.UTF8);
        writer.Write(text);
    }

    private static string Escape(string text) => text.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;");

    private static byte[] RewritePackage(byte[] bytes, Func<string, string, string> rewrite)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in zip.Entries.ToArray())
            {
                string name = entry.FullName;
                string text;
                using (var reader = new StreamReader(entry.Open()))
                    text = reader.ReadToEnd();
                entry.Delete();
                WriteEntry(zip, name, rewrite(name, text));
            }
        }

        return stream.ToArray();
    }

    private sealed class FailingScanStream(byte[] bytes) : MemoryStream(bytes)
    {
        public bool Fail { get; set; }

        public IOException Error { get; } = new("Hyperlink scan failed.");

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (Fail)
                throw Error;
            return base.Read(buffer, offset, count);
        }

#if NET8_0_OR_GREATER
        public override int Read(Span<byte> buffer)
        {
            if (Fail)
                throw Error;
            return base.Read(buffer);
        }
#endif
    }
}