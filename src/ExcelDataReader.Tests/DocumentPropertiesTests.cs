#nullable enable

using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using ExcelDataReader.Core;
using ExcelDataReader.Core.CompoundFormat;

namespace ExcelDataReader.Tests;

public sealed class DocumentPropertiesTests
{
    private static readonly string[] ExcelFiles =
    [
        "DocumentProperties.xls", "DocumentProperties.xlsx", "DocumentProperties.xlsb", "DocumentProperties.xml",
        "DocumentPropertiesPassword.xls", "DocumentPropertiesPassword.xlsx", "DocumentPropertiesPassword.xlsb",
    ];

    [Test]
    public void ExcelSavedPropertiesAreAvailableDuringOpenAndSurviveReading(
        [ValueSource(nameof(ExcelFiles))] string file,
        [Values(false, true)] bool singlePass)
    {
        using var source = Configuration.GetTestWorkbook(file);
        using var reader = ExcelReaderFactory.CreateReader(source, new ExcelReaderConfiguration
        {
            Password = "password",
            SinglePassMode = singlePass,
            LeaveOpen = true,
        });
        var properties = reader.DocumentProperties;
        Assert.That(properties.Title, Is.EqualTo("Workbook title"));
        Assert.That(properties.Subject, Is.EqualTo("Workbook subject"));
        Assert.That(properties.Keywords, Is.EqualTo("excel; test; keywords"));
        Assert.That(properties.Description, Is.EqualTo("Description & <text>\nSecond line"));

        // Excel's RemovePersonalInformation option removes both author fields on save.
        Assert.That(properties.Author, Is.Null.Or.Empty);
        Assert.That(properties.LastModifiedBy, Is.Null.Or.Empty);
        Assert.That(properties.Created, Is.EqualTo(new DateTime(2020, 1, 2, 2, 4, 5, DateTimeKind.Utc)));
        Assert.That(properties.LastPrinted, Is.EqualTo(new DateTime(2020, 5, 6, 5, 8, 8, DateTimeKind.Utc)));
        Assert.That(properties.Modified, Is.GreaterThan(properties.Created));
        Assert.That(properties.Created!.Value.Kind, Is.EqualTo(DateTimeKind.Utc));
        Assert.That(properties.Modified!.Value.Kind, Is.EqualTo(DateTimeKind.Utc));
        Assert.That(properties.LastPrinted!.Value.Kind, Is.EqualTo(DateTimeKind.Utc));
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Name, Is.EqualTo("Data"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo("Document properties fixture"));
            Assert.That(reader.GetDouble(1), Is.EqualTo(42));
            reader.ResetWorksheet();
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetDouble(1), Is.EqualTo(42));
            Assert.That(reader.Read(), Is.False);
            Assert.That(reader.NextResult(), Is.True);
            Assert.That(reader.Name, Is.EqualTo("Other"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetString(0), Is.EqualTo("Second sheet"));
            Assert.That(reader.Read(), Is.False);
            Assert.That(reader.NextResult(), Is.False);
            Assert.That(reader.DocumentProperties, Is.SameAs(properties));
            reader.Reset();
        }

        var dataSet = reader.AsDataSet();
        Assert.That(dataSet.Tables.Count, Is.EqualTo(2));
        Assert.That(dataSet.Tables[0].Rows[0][1], Is.EqualTo(42));
        reader.Close();
        Assert.That(source.CanRead, Is.True);
        Assert.That(reader.DocumentProperties, Is.SameAs(properties));
        Assert.That(properties.Title, Is.EqualTo("Workbook title"));
    }

    [TestCase("xls")]
    [TestCase("xlsx")]
    [TestCase("xlsb")]
    public void FormatSpecificFactoriesAlsoLoadProperties(string format)
    {
        using var source = Configuration.GetTestWorkbook("DocumentProperties." + format);
        using var reader = format == "xls"
            ? ExcelReaderFactory.CreateBinaryReader(source)
            : ExcelReaderFactory.CreateOpenXmlReader(source);
        Assert.That(reader.DocumentProperties.Title, Is.EqualTo("Workbook title"));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(1), Is.EqualTo(42));
    }

    [TestCase(1252, false)]
    [TestCase(1251, false)]
    [TestCase(1200, false)]
    [TestCase(65001, false)]
    [TestCase(1252, true)]
    public void BinaryPropertiesRespectCodePageUnicodeAndTerminators(int codePage, bool wideString)
    {
        string title = codePage == 1251 ? "\u043e\u0442\u0447\u0451\u0442" : "caf\u00e9";
        byte[] bytes = Summary(
            (1, Scalar(2, writer => writer.Write((ushort)codePage))),
            (2, Text(title, Encoding.GetEncoding(codePage), wideString)),
            (4, Text("\u5831\u544a", Encoding.Unicode, true)),
            (8, Text("Fixture Editor", Encoding.Unicode, true)),
            (12, Scalar(64, writer => writer.Write(new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc).ToFileTimeUtc()))));
        using var source = new MemoryStream(bytes);
        var properties = CompoundDocumentPropertiesReader.Read(source);
        Assert.That(properties.Title, Is.EqualTo(title));
        Assert.That(properties.Author, Is.EqualTo("\u5831\u544a"));
        Assert.That(properties.LastModifiedBy, Is.EqualTo("Fixture Editor"));
        Assert.That(properties.Created, Is.EqualTo(new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc)));
        Assert.That(source.CanRead, Is.True);
    }

    [TestCase(30)]
    [TestCase(31)]
    [TestCase(19)]
    public void InvalidOrUnsupportedBinaryPropertyDoesNotHideOtherProperties(int type)
    {
        byte[] bytes = Summary(
            (2, Scalar((uint)type, writer => writer.Write(uint.MaxValue))),
            (4, Text("Fixture Author", Encoding.Unicode, true)),
            (13, Scalar(64, writer => writer.Write(long.MaxValue))));
        using var source = new MemoryStream(bytes);
        var properties = CompoundDocumentPropertiesReader.Read(source);
        Assert.That(properties.Title, Is.Null);
        Assert.That(properties.Author, Is.EqualTo("Fixture Author"));
        Assert.That(properties.Modified, Is.Null);
    }

    [TestCase(0)]
    [TestCase(65535)]
    public void UnknownBinaryCodePageDoesNotHideUnicodeProperties(int codePage)
    {
        using var source = new MemoryStream(Summary(
            (1, Scalar(2, writer => writer.Write((ushort)codePage))),
            (2, Text("ANSI title", Encoding.ASCII, false)),
            (4, Text("Fixture Author", Encoding.Unicode, true))));
        var properties = CompoundDocumentPropertiesReader.Read(source);
        Assert.That(properties.Title, Is.Null);
        Assert.That(properties.Author, Is.EqualTo("Fixture Author"));
    }

    [Test]
    public void RawBiffRemainsReadableWithoutProperties()
    {
        using var source = Configuration.GetTestWorkbook("as3xls_BIFF2.xls");
        using var reader = ExcelReaderFactory.CreateReader(source);
        Assert.That(reader.DocumentProperties.Title, Is.Null);
        Assert.That(reader.Read(), Is.True);
    }

    [Test]
    public void EmptySummaryStreamDoesNotRequireAValidSectorChain()
    {
        using var source = Configuration.GetTestWorkbook("DocumentProperties.xls");
        var document = new CompoundDocument(source);
        var entry = document.FindEntry("\u0005SummaryInformation")!;
        entry.StreamSize = 0;
        entry.StreamFirstSector = uint.MaxValue;
        Assert.That(CompoundDocumentPropertiesReader.Load(document, source).Title, Is.Null);
    }

    [Test]
    public void PropertyReadersDoNotHideCallerIoErrors()
    {
        using var binary = new FailingStream(Summary((2, Text("Title", Encoding.Unicode, true))));
        Assert.That(Assert.Throws<IOException>(() => CompoundDocumentPropertiesReader.Read(binary)), Is.SameAs(binary.Error));
        using var xmlStream = new FailingStream(Encoding.UTF8.GetBytes("<coreProperties/>"));
        Assert.That(
            Assert.Throws<IOException>(() =>
            {
                using var xml = System.Xml.XmlReader.Create(xmlStream);
                new DocumentPropertiesBuilder().ReadXml(xml, spreadsheetXml: false);
            }),
            Is.SameAs(xmlStream.Error));
    }

    [TestCase(24)]
    [TestCase(44)]
    [TestCase(48)]
    [TestCase(52)]
    [TestCase(60)]
    public void InvalidBinarySectionBoundsAreIgnored(int offset)
    {
        byte[] bytes = Summary((2, Text("Title", Encoding.Unicode, true)));
        Array.Copy(BitConverter.GetBytes(uint.MaxValue), 0, bytes, offset, 4);
        using var source = new MemoryStream(bytes);
        Assert.That(CompoundDocumentPropertiesReader.Read(source).Title, Is.Null);
    }

    [TestCase("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"><")]
    [TestCase("<!DOCTYPE x SYSTEM \"file:///not-a-real-document\"><x/>")]
    [TestCase("<unexpected/>")]
    public void UnreadableOptionalCorePropertiesDoNotPreventOpening(string xml)
    {
        using var source = Package(xml);
        using var reader = ExcelReaderFactory.CreateReader(source);
        Assert.That(reader.DocumentProperties.Title, Is.Null);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(1), Is.EqualTo(42));
    }

    [Test]
    public void CorePropertiesUseNamespacesAndIgnoreUnsupportedValues()
    {
        const string xml = """
            <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
              xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:other="urn:other">
              <other:title>Wrong namespace</other:title>
              <dc:title>Title &amp; text</dc:title>
              <dc:creator>Fixture Author</dc:creator>
              <cp:lastModifiedBy>Fixture Editor</cp:lastModifiedBy>
              <dc:subject><other:value>Not simple text</other:value></dc:subject>
              <dcterms:created>2020-01-02T03:04:05+02:00</dcterms:created>
              <dcterms:modified>invalid date</dcterms:modified>
              <cp:unknown>Ignored</cp:unknown>
            </cp:coreProperties>
            """;
        using var reader = ExcelReaderFactory.CreateReader(Package(xml));
        Assert.That(reader.DocumentProperties.Title, Is.EqualTo("Title & text"));
        Assert.That(reader.DocumentProperties.Author, Is.EqualTo("Fixture Author"));
        Assert.That(reader.DocumentProperties.LastModifiedBy, Is.EqualTo("Fixture Editor"));
        Assert.That(reader.DocumentProperties.Subject, Is.Null);
        Assert.That(reader.DocumentProperties.Created, Is.EqualTo(new DateTime(2020, 1, 2, 1, 4, 5, DateTimeKind.Utc)));
        Assert.That(reader.DocumentProperties.Modified, Is.Null);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void CorePropertiesFollowPackageRelationshipRatherThanFixedPath(bool external)
    {
        using var source = Package(null);
        using (var zip = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("docProps/core.xml")!;
            byte[] properties;
            using (var buffer = new MemoryStream())
            {
                using (var input = entry.Open())
                    input.CopyTo(buffer);
                properties = buffer.ToArray();
            }

            entry.Delete();
            using (var output = zip.CreateEntry("metadata/core properties.xml").Open())
                output.Write(properties, 0, properties.Length);
            var relationships = zip.GetEntry("_rels/.rels")!;
            XDocument xml;
            using (var input = relationships.Open())
                xml = XDocument.Load(input);
            var relation = xml.Root!.Elements().Single(element => ((string?)element.Attribute("Type"))?.EndsWith("/metadata/core-properties", StringComparison.Ordinal) == true);
            relation.SetAttributeValue("Target", "/metadata/core%20properties.xml");
            if (external)
                relation.SetAttributeValue("TargetMode", "External");
            relationships.Delete();
            using var outputXml = zip.CreateEntry("_rels/.rels").Open();
            xml.Save(outputXml);
        }

        source.Position = 0;
        using var reader = ExcelReaderFactory.CreateReader(source);
        Assert.That(reader.DocumentProperties.Title, external ? Is.Null : Is.EqualTo("Workbook title"));
        Assert.That(reader.Read(), Is.True);
    }

    [Test]
    public void SpreadsheetXmlReadsAuthorsAndSkipsInvalidDates()
    {
        const string xml = """
            <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
              <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
                <Author>Fixture Author</Author><LastAuthor>Fixture Editor</LastAuthor>
                <Created>2020-01-02T03:04:05Z</Created><LastSaved>not a date</LastSaved>
              </DocumentProperties>
              <Worksheet Name="Data"><Table><Row><Cell><Data Type="Number">42</Data></Cell></Row></Table></Worksheet>
            </Workbook>
            """;
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(Encoding.UTF8.GetBytes(xml)));
        Assert.That(reader.DocumentProperties.Author, Is.EqualTo("Fixture Author"));
        Assert.That(reader.DocumentProperties.LastModifiedBy, Is.EqualTo("Fixture Editor"));
        Assert.That(reader.DocumentProperties.Created, Is.EqualTo(new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc)));
        Assert.That(reader.DocumentProperties.Modified, Is.Null);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(42));
    }

    [Test]
    public void MissingCorePartAndCsvReturnEmptyProperties()
    {
        using var source = Package(null);
        using (var zip = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true))
            zip.GetEntry("docProps/core.xml")!.Delete();
        source.Position = 0;
        using var reader = ExcelReaderFactory.CreateReader(source);
        Assert.That(reader.DocumentProperties.Title, Is.Null);
        Assert.That(reader.Read(), Is.True);
        using var csv = ExcelReaderFactory.CreateCsvReader(new MemoryStream(Encoding.UTF8.GetBytes("a,b")));
        Assert.That(csv.DocumentProperties.Title, Is.Null);
        Assert.That(csv.Read(), Is.True);
        Assert.That(csv.GetString(0), Is.EqualTo("a"));
    }

    private static MemoryStream Package(string? coreXml)
    {
        var stream = new MemoryStream();
        using (var source = Configuration.GetTestWorkbook("DocumentProperties.xlsx"))
            source.CopyTo(stream);
        if (coreXml != null)
        {
            using var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true);
            zip.GetEntry("docProps/core.xml")!.Delete();
            using var writer = new StreamWriter(zip.CreateEntry("docProps/core.xml").Open());
            writer.Write(coreXml);
        }

        stream.Position = 0;
        return stream;
    }

    private static byte[] Scalar(uint type, Action<BinaryWriter> write)
    {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream);
        writer.Write(type);
        write(writer);
        return stream.ToArray();
    }

    private static byte[] Text(string value, Encoding encoding, bool wide) => Scalar(wide ? 31U : 30U, writer =>
    {
        byte[] text = (wide ? Encoding.Unicode : encoding).GetBytes(value + "\0");
        writer.Write((uint)(wide ? text.Length / 2 : text.Length));
        writer.Write(text);
        writer.Write(new byte[3]);
    });

    private static byte[] Summary(params (uint Id, byte[] Value)[] properties)
    {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream);
        writer.Write((ushort)0xFFFE);
        writer.Write(new byte[22]);
        writer.Write(1U);
        writer.Write(new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9").ToByteArray());
        writer.Write(48U);
        writer.Write((uint)(8 + properties.Length * 8 + properties.Sum(property => property.Value.Length)));
        writer.Write((uint)properties.Length);
        uint offset = (uint)(8 + properties.Length * 8);
        foreach (var property in properties)
        {
            writer.Write(property.Id);
            writer.Write(offset);
            offset += (uint)property.Value.Length;
        }

        foreach (var property in properties)
            writer.Write(property.Value);
        return stream.ToArray();
    }

    private sealed class FailingStream(byte[] bytes) : MemoryStream(bytes)
    {
        public IOException Error { get; } = new("Caller read failure");

        public override int Read(byte[] buffer, int offset, int count) => throw Error;

        public override int ReadByte() => throw Error;

#if NET8_0_OR_GREATER
        public override int Read(Span<byte> buffer) => throw Error;
#endif
    }
}