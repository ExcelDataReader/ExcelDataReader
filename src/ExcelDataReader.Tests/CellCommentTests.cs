#nullable enable

using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Tests;

public sealed class CellCommentTests
{
    private const string Author = "Author \u03bb";
    private const string Text = "first & <second>\n\u03bb \ud83d\ude00";
    private const int Rows = 700;

    [TestCase("xml", false)]
    [TestCase("xml", true)]
    [TestCase("xls", false)]
    [TestCase("xls", true)]
    [TestCase("xlsx", false)]
    [TestCase("xlsx", true)]
    [TestCase("xlsb", false)]
    [TestCase("xlsb", true)]
    public void LazyCommentsPreserveBufferedRowsAndReaderLifecycle(string format, bool singlePass)
    {
        byte[] bytes = CreateWorkbook(format);
        using var source = new MemoryStream(bytes);
        using var reader = ExcelReaderFactory.CreateReader(source, new ExcelReaderConfiguration { SinglePassMode = singlePass, LeaveOpen = true });
        using var baseline = ExcelReaderFactory.CreateReader(new MemoryStream(bytes), new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.Throws<InvalidOperationException>(() => reader.GetComment(0));
        CellCommentThread? retained = null;
        for (int row = 0; row < Rows; row++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(baseline.Read(), Is.True);
            Assert.That(reader.GetValue(0), Is.EqualTo(baseline.GetValue(0)), $"row {row}");
            long position = source.Position;
            if (row == 13)
            {
                retained = reader.GetComment(0);
                AssertMessage(retained, Author, Text);
                Assert.That(reader.GetComment(0), Is.SameAs(retained));
                Assert.That(reader.GetValue(0), Is.EqualTo(baseline.GetValue(0)));
                Assert.Throws<IndexOutOfRangeException>(() => reader.GetComment(-1));
                Assert.Throws<IndexOutOfRangeException>(() => reader.GetComment(reader.FieldCount));
            }
            else if (row > 13)
            {
                Assert.That(reader.GetComment(0), Is.Null);
            }

            if (format is "xml" or "xls")
                Assert.That(source.Position, Is.EqualTo(position), $"physical stream position at row {row}");
        }

        Assert.That(reader.Read(), Is.False);
        Assert.Throws<InvalidOperationException>(() => reader.GetComment(0));
        reader.Reset();
        for (int attempt = 0; attempt < 2; attempt++)
        {
            for (int row = 0; row <= 13; row++)
                Assert.That(reader.Read(), Is.True);
            AssertMessage(reader.GetComment(0), Author, Text);
            Assert.That(reader.GetComment(0), Is.SameAs(retained));
            reader.ResetWorksheet();
        }

        reader.Close();
        Assert.That(source.CanRead, Is.True);
        Assert.Throws<InvalidOperationException>(() => reader.GetComment(0));
        AssertMessage(retained, Author, Text);
    }

    [TestCase("Issue242_StdRc4PwdPassword.xls")]
    [TestCase("Issue242_XorPwdPassword.xls")]
    [TestCase("agile_AES128_SHA1_CBC_pwd_password.xlsx")]
    [TestCase("agile_AES128_SHA1_CBC_pwd_password.xlsb")]
    [TestCase("standard_AES128_SHA1_ECB_pwd_password.xlsx")]
    public void AuxiliaryScanPreservesEncryptedTransport(string file)
    {
        using var source = Configuration.GetTestWorkbook(file);
        using var reader = ExcelReaderFactory.CreateReader(source, new ExcelReaderConfiguration { Password = "password" });
        using var baseline = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook(file), new ExcelReaderConfiguration { Password = "password" });
        bool hasResult;
        do
        {
            while (reader.Read())
            {
                Assert.That(baseline.Read(), Is.True);
                long position = source.Position;
                if (reader.FieldCount > 0)
                    _ = reader.GetComment(0);
                if (file.EndsWith(".xls", StringComparison.Ordinal))
                    Assert.That(source.Position, Is.EqualTo(position));
                for (int column = 0; column < reader.FieldCount; column++)
                    Assert.That(reader.GetValue(column), Is.EqualTo(baseline.GetValue(column)));
            }

            Assert.That(baseline.Read(), Is.False);
            hasResult = reader.NextResult();
            Assert.That(hasResult, Is.EqualTo(baseline.NextResult()));
        }
        while (hasResult);
    }

    [TestCase(2)]
    [TestCase(3)]
    [TestCase(4)]
    [TestCase(5)]
    public void LegacyBiffNotesIncludeContinuedText(int version)
    {
        string text = new('x', 3000);
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(CreateXls(version, text)));
        for (int row = 0; row <= 13; row++)
            Assert.That(reader.Read(), Is.True);
        AssertMessage(reader.GetComment(0), null, text);
    }

    [TestCase(2)]
    [TestCase(3)]
    [TestCase(4)]
    [TestCase(5)]
    public void LegacyBiffRejectsInvalidContinuationLengths(int version)
    {
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(CreateXls(version, new string('x', 3000), malformed: true)));
        Assert.That(reader.Read(), Is.True);
        Assert.Throws<ExcelReaderException>(() => reader.GetComment(0));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(1));
    }

    [Test]
    public void CsvUsesNormalAccessorValidationButHasNoComments()
    {
        using var reader = ExcelReaderFactory.CreateCsvReader(new MemoryStream(Encoding.UTF8.GetBytes("a,b\nc,d")));
        Assert.Throws<InvalidOperationException>(() => reader.GetComment(0));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetComment(0), Is.Null);
        Assert.Throws<IndexOutOfRangeException>(() => reader.GetComment(2));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.Read(), Is.False);
        Assert.Throws<InvalidOperationException>(() => reader.GetComment(0));
    }

    [TestCase("xml")]
    [TestCase("xlsx")]
    [TestCase("xlsb")]
    [TestCase("xls")]
    public void CommentsAreNotDecodedUnlessRequestedAndFailuresRestoreStream(string format)
    {
        byte[] bytes = CreateWorkbook(format, malformed: true);
        using var source = new MemoryStream(bytes);
        using var reader = ExcelReaderFactory.CreateReader(source);
        for (int row = 0; row <= 13; row++)
            Assert.That(reader.Read(), Is.True);
        long position = source.Position;
        for (int attempt = 0; attempt < 2; attempt++)
        {
            Assert.Throws<ExcelReaderException>(() => reader.GetComment(0));
            if (format is "xml" or "xls")
                Assert.That(source.Position, Is.EqualTo(position));
        }

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(14));
        while (reader.Read())
        {
        }
    }

    [TestCase("Issue769_CommentAndTextInA1.xml")]
    [TestCase("Issue769_TextInA1-CommentInB1.xml")]
    public void RealSpreadsheetMlNotesDoNotBecomeValues(string file)
    {
        using var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook(file));
        Assert.That(reader.Read(), Is.True);
        int column = file.Contains("B1") ? 1 : 0;
        if (column == 1)
        {
            Assert.That(reader.FieldCount, Is.EqualTo(1));
            Assert.Throws<IndexOutOfRangeException>(() => reader.GetComment(1));
        }
        else
        {
            AssertMessage(reader.GetComment(column), "Johan Appelgren", "Johan Appelgren:\nLorem ipsum\n");
        }
    }

    [TestCase("xls")]
    [TestCase("xlsx")]
    [TestCase("xlsb")]
    [TestCase("xml")]
    public void RealRichTextCommentHasAuthorAndPlainText(string format)
    {
        using var reader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("OldIssue11573_BlankValues." + format));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.Read(), Is.True);
        AssertMessage(reader.GetComment(2), "MIRIAM", "\nPOLA SZARE WYPE\u0141NIA OSOBA ZAMAWIAJ\u0104CA");
    }

    [TestCase("")]
    [TestCase("first")]
    [TestCase("long")]
    [TestCase("surrogates")]
    public void Biff8SupportsEmptyAndMixedWidthContinuedNotes(string kind)
    {
        string text = kind switch
        {
            "long" => new string('x', 4100) + "\u03bb" + new string('y', 5000),
            "surrogates" => new string('x', 2047) + "\ud83d\ude00" + new string('y', 3000),
            _ => kind,
        };
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(CreateXls(8, text)));
        for (int row = 0; row <= 13; row++)
            Assert.That(reader.Read(), Is.True);
        AssertMessage(reader.GetComment(0), Author, text);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(14));
    }

    [TestCase(false)]
    [TestCase(true)]
    public void SpreadsheetMlBlankNotesGapsAndResultChanges(bool singlePass)
    {
        const string xml = """
            <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
                      xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
              <Worksheet ss:Name="First"><Table>
                <Row ss:Index="2"><Cell ss:MergeAcross="1">
                  <Comment ss:Author="one">
                    <ss:Data> before <b xmlns="http://www.w3.org/TR/REC-html40">bold</b> after </ss:Data>
                  </Comment>
                </Cell><Cell><Data ss:Type="Number">2</Data></Cell></Row>
                <Row ss:Index="4"><Cell><Comment><ss:Data/></Comment></Cell>
                  <Cell ss:Index="3"><Data ss:Type="Number">4</Data></Cell></Row>
              </Table></Worksheet>
              <Worksheet ss:Name="Second"><Table>
                <Row><Cell><Comment ss:Author="two"><ss:Data>second</ss:Data></Comment>
                  <Data ss:Type="String">value</Data></Cell></Row>
              </Table></Worksheet>
            </Workbook>
            """;
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(Encoding.UTF8.GetBytes(xml)), new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.IsDBNull(0), Is.True);
        AssertMessage(reader.GetComment(0), "one", " before bold after ");
        Assert.That(reader.GetComment(1), Is.Null);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetComment(0), Is.Null);
        Assert.That(reader.Read(), Is.True);
        AssertMessage(reader.GetComment(0), null, string.Empty);
        Assert.That(reader.NextResult(), Is.True);
        Assert.Throws<InvalidOperationException>(() => reader.GetComment(0));
        Assert.That(reader.Read(), Is.True);
        AssertMessage(reader.GetComment(0), "two", "second");
        reader.Reset();
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.Read(), Is.True);
        AssertMessage(reader.GetComment(0), "one", " before bold after ");
    }

    [TestCase(false)]
    [TestCase(true)]
    public void OpenXmlResolvesStrictAbsolutePathsAndIgnoresPhonetics(bool external)
    {
        byte[] bytes = CreateWorkbook("xlsx");
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            zip.GetEntry("xl/worksheets/_rels/data.xml.rels")!.Delete();
            WriteEntry(zip, "xl/worksheets/_rels/data.xml.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"notes\" Type=\"http://purl.oclc.org/ooxml/officeDocument/relationships/comments\" Target=\"{(external ? "https://example.invalid/comments.xml" : "/notes/comments.xml")}\"{(external ? " TargetMode=\"External\"" : string.Empty)}/></Relationships>");
            zip.GetEntry("notes/comments.xml")!.Delete();
            WriteEntry(zip, "notes/comments.xml", """
                <comments xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">
                  <authors><author>_x0041_uthor</author></authors>
                  <commentList>
                    <comment ref="A14" authorId="0"><text><t xml:space="preserve"> text </t>
                      <rPh sb="0" eb="1"><t>phonetic</t></rPh></text></comment>
                    <comment ref="A15" authorId="0"><text/></comment>
                  </commentList>
                </comments>
                """);
        }

        stream.Position = 0;
        using var reader = ExcelReaderFactory.CreateReader(stream);
        for (int row = 0; row <= 13; row++)
            Assert.That(reader.Read(), Is.True);
        if (external)
        {
            Assert.That(reader.GetComment(0), Is.Null);
        }
        else
        {
            AssertMessage(reader.GetComment(0), "Author", " text ");
            Assert.That(reader.Read(), Is.True);
            AssertMessage(reader.GetComment(0), "Author", string.Empty);
        }
    }

    [TestCase("A0", "0")]
    [TestCase("XFE1", "0")]
    [TestCase("A1048577", "0")]
    [TestCase("AAAAAAA1", "0")]
    [TestCase("A14", null)]
    [TestCase("A14", "-1")]
    [TestCase("A14", "1")]
    [TestCase("A14", "invalid")]
    [TestCase("A14", "2147483648")]
    public void InvalidOpenXmlCommentReferencesFailOnlyOnLookup(string reference, string? author)
    {
        using var stream = new MemoryStream();
        byte[] bytes = CreateWorkbook("xlsx");
        stream.Write(bytes, 0, bytes.Length);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            zip.GetEntry("notes/comments.xml")!.Delete();
            string attribute = author == null ? string.Empty : $" authorId=\"{author}\"";
            WriteEntry(zip, "notes/comments.xml", $"<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors><author>author</author></authors><commentList><comment ref=\"{reference}\"{attribute}><text><t>note</t></text></comment></commentList></comments>");
        }

        stream.Position = 0;
        using var reader = ExcelReaderFactory.CreateReader(stream);
        Assert.That(reader.Read(), Is.True);
        Assert.Throws<ExcelReaderException>(() => reader.GetComment(0));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(1));
    }

    [TestCase("missing-target")]
    [TestCase("missing-part")]
    [TestCase("outside-package")]
    [TestCase("wrong-root")]
    [TestCase("xml-syntax")]
    public void InvalidCommentPartsPreserveLazyFailureAndContinuation(string failure)
    {
        using var stream = new MemoryStream();
        byte[] bytes = CreateWorkbook("xlsx");
        stream.Write(bytes, 0, bytes.Length);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            if (failure is "wrong-root" or "xml-syntax")
            {
                zip.GetEntry("notes/comments.xml")!.Delete();
                WriteEntry(zip, "notes/comments.xml", failure == "wrong-root" ? "<wrong/>" : "<comments");
            }
            else
            {
                zip.GetEntry("xl/worksheets/_rels/data.xml.rels")!.Delete();
                string target = failure switch
                {
                    "missing-target" => string.Empty,
                    "missing-part" => " Target=\"../../notes/missing.xml\"",
                    _ => " Target=\"../../../comments.xml\"",
                };
                WriteEntry(zip, "xl/worksheets/_rels/data.xml.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"notes\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\"{target}/></Relationships>");
            }
        }

        stream.Position = 0;
        using var reader = ExcelReaderFactory.CreateReader(stream);
        Assert.That(reader.Read(), Is.True);
        for (int attempt = 0; attempt < 2; attempt++)
        {
            if (failure == "xml-syntax")
                Assert.Throws<XmlException>(() => reader.GetComment(0));
            else
                Assert.Throws<ExcelReaderException>(() => reader.GetComment(0));
        }

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(1));
        reader.Reset();
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.Zero);
    }

    [TestCase("xml", false)]
    [TestCase("xml", true)]
    [TestCase("xls", false)]
    [TestCase("xls", true)]
    public void CommentRestorationFailuresPreserveScanErrors(string format, bool failScan)
    {
        using var source = new CommentFailureStream(CreateWorkbook(format));
        using var reader = ExcelReaderFactory.CreateReader(source, new ExcelReaderConfiguration { LeaveOpen = true });
        Assert.That(reader.Read(), Is.True);
        source.RestorePosition = source.Position;
        source.FailScan = failScan;
        source.Armed = true;
        try
        {
            if (failScan)
            {
                Assert.That(Assert.Throws<IOException>(() => reader.GetComment(0)), Is.SameAs(source.ScanError));
            }
            else
            {
                Assert.That(Assert.Throws<IOException>(() => reader.GetComment(0)), Is.SameAs(source.RestoreError));
            }
        }
        finally
        {
            source.Armed = false;
        }
    }

    [TestCase("xlsx", false)]
    [TestCase("xlsx", true)]
    [TestCase("xls", false)]
    [TestCase("xls", true)]
    public void CommentLoadingDoesNotTranslateCallerStreamExceptions(string format, bool xmlError)
    {
        Exception failure = xmlError ? new XmlException("Caller stream failure.") : new ArgumentException("Caller stream failure.");
        using var source = new CommentFailureStream(CreateWorkbook(format)) { ScanError = failure };
        using var reader = ExcelReaderFactory.CreateReader(source, new ExcelReaderConfiguration { LeaveOpen = true });
        Assert.That(reader.Read(), Is.True);
        source.RestorePosition = long.MaxValue;
        source.FailScan = true;
        source.Armed = true;
        try
        {
            Assert.That(Assert.Throws(failure.GetType(), () => reader.GetComment(0)), Is.SameAs(failure));
        }
        finally
        {
            source.Armed = false;
        }

        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(1));
    }

    private static void AssertMessage(CellCommentThread? comment, string? author, string text)
    {
        Assert.That(comment, Is.Not.Null);
        Assert.That(comment!.Messages, Has.Count.EqualTo(1));
        Assert.That(comment.Messages[0].Author, Is.EqualTo(author));
        Assert.That(comment.Messages[0].Text, Is.EqualTo(text));
    }

    private static byte[] CreateWorkbook(string format, bool malformed = false)
    {
        if (format == "xls")
            return CreateXls(8, Text, malformed);
        if (format == "xml")
        {
            var xml = new StringBuilder("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"><Worksheet ss:Name=\"Notes\"><Table>");
            for (int row = 0; row < Rows; row++)
            {
                xml.AppendFormat(CultureInfo.InvariantCulture, "<Row><Cell><Data ss:Type=\"Number\">{0}</Data>", row);
                if (row == 13)
                    xml.Append($"<Comment ss:Author=\"{Author}\"><ss:Data>{Escape(Text)}</ss:Data></Comment>");
                if (malformed && row == 13)
                    xml.Append("<Comment><ss:Data>duplicate</ss:Data></Comment>");
                xml.Append("</Cell></Row>");
            }

            xml.Append("</Table></Worksheet></Workbook>");
            return Encoding.UTF8.GetBytes(xml.ToString());
        }

        bool binary = format == "xlsb";
        string extension = binary ? "bin" : "xml";
        using var buffer = new MemoryStream();
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(zip, $"xl/_rels/workbook.{extension}.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"sheet\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/data.{extension}\"/></Relationships>");
            WriteEntry(zip, $"xl/worksheets/_rels/data.{extension}.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"notes\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../../notes/./comments.{extension}\"/></Relationships>");
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
                        WideString(data, "Notes");
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
                            data.Write((byte)0);
                            data.Write((byte)0);
                        });
                        Record(writer, 5, data =>
                        {
                            data.Write(0);
                            data.Write(0);
                            data.Write((double)row);
                        });
                    }

                    Record(writer, 0x92, _ => { });
                }

                using (var writer = new BinaryWriter(zip.CreateEntry("notes/comments.bin").Open()))
                {
                    Record(writer, 0x274, _ => { });
                    Record(writer, 0x276, _ => { });
                    Record(writer, 0x278, data => WideString(data, Author));
                    Record(writer, 0x277, _ => { });
                    Record(writer, 0x279, _ => { });
                    Record(writer, 0x27B, data =>
                    {
                        data.Write(malformed ? 99 : 0);
                        data.Write(13);
                        data.Write(13);
                        data.Write(0);
                        data.Write(0);
                        data.Write(new byte[16]);
                    });
                    Record(writer, 0x27D, data =>
                    {
                        data.Write((byte)1);
                        WideString(data, Text);
                        data.Write(0);
                    });
                    Record(writer, 0x27C, _ => { });
                    Record(writer, 0x27A, _ => { });
                    Record(writer, 0x275, _ => { });
                }
            }
            else
            {
                WriteEntry(zip, "xl/workbook.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Notes\" sheetId=\"1\" r:id=\"sheet\"/></sheets></workbook>");
                var xml = new StringBuilder("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
                for (int row = 0; row < Rows; row++)
                    xml.AppendFormat(CultureInfo.InvariantCulture, "<row r=\"{0}\"><c r=\"A{0}\"><v>{1}</v></c></row>", row + 1, row);
                xml.Append("</sheetData></worksheet>");
                WriteEntry(zip, "xl/worksheets/data.xml", xml.ToString());
                WriteEntry(zip, "notes/comments.xml", $"<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors><author>{Author}</author></authors><commentList><comment ref=\"A14\" authorId=\"{(malformed ? 99 : 0)}\"><text><r><t>first &amp; </t></r><r><t>&lt;second&gt;\n\u03bb \ud83d\ude00</t></r></text></comment></commentList></comments>");
            }
        }

        return buffer.ToArray();
    }

    private static byte[] CreateXls(int version, string text, bool malformed = false)
    {
        using var buffer = new MemoryStream();
        using var writer = new BinaryWriter(buffer);
        ushort bof = version switch { 2 => 9, 3 => 0x209, 4 => 0x409, _ => 0x809 };
        XlsRecord(writer, bof, data =>
        {
            data.Write((ushort)(version == 8 ? 0x600 : 0x500));
            data.Write((ushort)0x10);
            data.Write(new byte[version == 2 ? 0 : version < 5 ? 2 : 12]);
        });
        for (int row = 0; row < Rows; row++)
        {
            XlsRecord(writer, version == 2 ? (ushort)3 : (ushort)0x203, data =>
            {
                data.Write((ushort)row);
                data.Write((ushort)0);
                data.Write((ushort)0);
                if (version == 2)
                    data.Write((byte)0);
                data.Write((double)row);
            });
        }

        if (version == 8)
        {
            XlsRecord(writer, 0x5D, data =>
            {
                data.Write((ushort)0x15);
                data.Write((ushort)18);
                data.Write((ushort)0x19);
                data.Write((ushort)7);
                data.Write(new byte[14]);
                data.Write(0);
            });
            XlsRecord(writer, 0x1B6, data =>
            {
                data.Write(new byte[10]);
                data.Write((ushort)text.Length);
                data.Write((ushort)(text.Length == 0 ? 0 : 16));
                data.Write((ushort)0);
                data.Write((ushort)0);
            });
            for (int offset = 0; offset < text.Length; offset += 2048)
            {
                string part = text.Substring(offset, Math.Min(2048, text.Length - offset));
                bool wide = part.Any(character => character > 255);
                XlsRecord(writer, 0x3C, data =>
                {
                    data.Write((byte)(wide ? 1 : 0));
                    if (wide)
                    {
                        foreach (char character in part)
                            data.Write((ushort)character);
                    }
                    else
                    {
                        data.Write(part.Select(character => (byte)character).ToArray());
                    }
                });
            }

            if (text.Length > 0)
                XlsRecord(writer, 0x3C, data => data.Write(new byte[16]));
            XlsRecord(writer, 0x1C, data =>
            {
                data.Write((ushort)13);
                data.Write((ushort)0);
                data.Write((ushort)0);
                data.Write((ushort)(malformed ? 99 : 7));
                data.Write((ushort)Author.Length);
                data.Write((byte)1);
                data.Write(Encoding.Unicode.GetBytes(Author));
                data.Write((byte)0);
            });
        }
        else
        {
            XlsRecord(writer, 0x42, data => data.Write((ushort)1252));
            for (int offset = 0; offset < text.Length; offset += 2048)
            {
                XlsRecord(writer, 0x1C, data =>
                {
                    data.Write((ushort)(offset == 0 ? 13 : ushort.MaxValue));
                    data.Write((ushort)0);
                    data.Write((ushort)(offset == 0 ? text.Length : malformed ? 0 : Math.Min(2048, text.Length - offset)));
                    data.Write(Encoding.ASCII.GetBytes(text.Substring(offset, Math.Min(2048, text.Length - offset))));
                });
            }
        }

        XlsRecord(writer, 0xA, _ => { });
        return buffer.ToArray();
    }

    private static void XlsRecord(BinaryWriter writer, ushort id, Action<BinaryWriter> write)
    {
        using var payload = new MemoryStream();
        using (var data = new BinaryWriter(payload, Encoding.Unicode, leaveOpen: true))
            write(data);
        writer.Write(id);
        writer.Write((ushort)payload.Length);
        writer.Write(payload.ToArray());
    }

    private static void Record(BinaryWriter writer, uint id, Action<BinaryWriter> write)
    {
        using var payload = new MemoryStream();
        using (var data = new BinaryWriter(payload, Encoding.Unicode, leaveOpen: true))
            write(data);
        WriteVariable(writer, id);
        WriteVariable(writer, (uint)payload.Length);
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

    private static void WideString(BinaryWriter writer, string value)
    {
        writer.Write(value.Length);
        writer.Write(Encoding.Unicode.GetBytes(value));
    }

    private static void WriteEntry(ZipArchive zip, string name, string value)
    {
        using var writer = new StreamWriter(zip.CreateEntry(name).Open(), Encoding.UTF8);
        writer.Write(value);
    }

    private static string Escape(string text)
    {
        var buffer = new StringBuilder();
        using (var writer = XmlWriter.Create(buffer, new XmlWriterSettings { ConformanceLevel = ConformanceLevel.Fragment }))
            writer.WriteString(text);
        return buffer.ToString();
    }

    private sealed class CommentFailureStream(byte[] bytes) : MemoryStream(bytes)
    {
        public bool Armed { get; set; }

        public bool FailScan { get; set; }

        public long RestorePosition { get; set; }

        public Exception ScanError { get; set; } = new IOException("Comment scan failed.");

        public IOException RestoreError { get; } = new("Position restoration failed.");

        public override long Position
        {
            get
            {
                return base.Position;
            }

            set
            {
                if (Armed && value == RestorePosition)
                    throw RestoreError;
                base.Position = value;
            }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (Armed && FailScan)
                throw ScanError;
            return base.Read(buffer, offset, count);
        }

#if NET8_0_OR_GREATER
        public override int Read(Span<byte> buffer)
        {
            if (Armed && FailScan)
                throw ScanError;
            return base.Read(buffer);
        }
#endif

        public override int ReadByte()
        {
            if (Armed && FailScan)
                throw ScanError;
            return base.ReadByte();
        }
    }
}