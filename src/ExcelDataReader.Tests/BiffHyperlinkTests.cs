#nullable enable

using System.IO.Compression;
using System.Text;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Tests;

public sealed class BiffHyperlinkTests
{
    [TestCase("", "Sheet1!A1", "", "#Sheet1!A1")]
    [TestCase("", "", "Only a tooltip", "")]
    [TestCase(null, "Sheet1!A1", "", "#Sheet1!A1")]
    [TestCase("missing", "Sheet1!A1", "Not internal", null)]
    public void EmptyRelationshipIdRepresentsCurrentWorkbook(string? relationshipId, string location, string tooltip, string? expected)
    {
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(CreateWorkbook(relationshipId, location, tooltip)));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(42));
        var hyperlink = reader.GetHyperlink(0);
        if (expected == null)
        {
            Assert.That(hyperlink, Is.Null);
        }
        else
        {
            Assert.That(hyperlink, Is.Not.Null);
            Assert.That(hyperlink!.Address, Is.EqualTo(expected));
            Assert.That(hyperlink.Tooltip, Is.EqualTo(tooltip.Length == 0 ? null : tooltip));
        }
    }

    [TestCase(256, false)]
    [TestCase(4, true)]
    public void InvalidRelationshipIdFailsWhenRequested(int length, bool embeddedNull)
    {
        string relationshipId = embeddedNull ? "r\0Id" : new string('r', length);
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(CreateWorkbook(relationshipId, "Sheet1!A1", string.Empty)));
        Assert.That(reader.Read(), Is.True);
        Assert.Throws<ExcelReaderException>(() => reader.GetHyperlink(0));
        Assert.Throws<ExcelReaderException>(() => reader.GetHyperlink(0));
    }

    private static byte[] CreateWorkbook(string? relationshipId, string location, string tooltip)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            using (var writer = new StreamWriter(zip.CreateEntry("xl/_rels/workbook.bin.rels").Open()))
                writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"sheet\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.bin\"/></Relationships>");
            using (var writer = new BinaryWriter(zip.CreateEntry("xl/workbook.bin").Open()))
            {
                Record(writer, 0x83, _ => { });
                Record(writer, 0x9C, data =>
                {
                    data.Write(0);
                    data.Write(1);
                    WideString(data, "sheet");
                    WideString(data, "Sheet1");
                });
                Record(writer, 0x84, _ => { });
            }

            using (var writer = new BinaryWriter(zip.CreateEntry("xl/worksheets/sheet1.bin").Open()))
            {
                Record(writer, 0x91, _ => { });
                Record(writer, 0, data =>
                {
                    data.Write(0);
                    data.Write(0);
                    data.Write((ushort)300);
                    data.Write((ushort)0);
                });
                Record(writer, 5, data =>
                {
                    data.Write(0);
                    data.Write(0);
                    data.Write(42.0);
                });
                Record(writer, 0x92, _ => { });
                Record(writer, 494, data =>
                {
                    data.Write(new byte[16]);
                    WideString(data, relationshipId);
                    WideString(data, location);
                    WideString(data, tooltip);
                    WideString(data, "Display label");
                });
            }
        }

        return output.ToArray();
    }

    private static void WideString(BinaryWriter writer, string? value)
    {
        writer.Write(value == null ? uint.MaxValue : (uint)value.Length);
        if (value != null)
            writer.Write(Encoding.Unicode.GetBytes(value));
    }

    private static void Record(BinaryWriter writer, uint id, Action<BinaryWriter> write)
    {
        using var payload = new MemoryStream();
        using (var data = new BinaryWriter(payload, Encoding.Unicode, leaveOpen: true))
            write(data);
        Variable(writer, id);
        Variable(writer, (uint)payload.Length);
        writer.Write(payload.ToArray());
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
}