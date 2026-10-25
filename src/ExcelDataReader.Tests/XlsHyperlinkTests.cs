#nullable enable

using System.Text;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Tests;

public sealed class XlsHyperlinkTests
{
    private const string FileMoniker = "00000303-0000-0000-C000-000000000046";
    private const string UrlMoniker = "79EAC9E0-BAF9-11CE-8C82-00AA004BA90B";

    [Test]
    public void UrlMonikerConsumesOptionalTailBeforeLocation([Values] bool hasTail, [Values] bool singlePass)
    {
        const string url = "https://example.com/a%23b#first";
        byte[] bytes = CreateWorkbook(writer => Hyperlink(
            writer,
            0,
            UrlMoniker,
            data =>
            {
                byte[] text = Encoding.Unicode.GetBytes(url + "\0");
                data.Write(text.Length + (hasTail ? 24 : 0));
                data.Write(text);
                if (hasTail)
                {
                    data.Write(new Guid("F4815879-1D3B-487F-AF2C-825DC4852763").ToByteArray());
                    data.Write(0);
                    data.Write(0);
                }
            },
            "second"));
        using var reader = Open(bytes, singlePass);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetHyperlink(0)?.Address, Is.EqualTo(url + "#second"));
        Assert.That(reader.GetDouble(0), Is.Zero);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(10));
    }

    [TestCase("Reports\\Budget.xlsx", 2, null, "..\\..\\Reports\\Budget.xlsx")]
    [TestCase("..\\Reports\\Budget.xlsx", 1, null, "..\\Reports\\Budget.xlsx")]
    [TestCase("caf\u00e9.xls", 0, null, "caf\u00e9.xls")]
    [TestCase("Reports\\?.xlsx", 1, "..\\Reports\\\u5831\u544a.xlsx", "..\\Reports\\\u5831\u544a.xlsx")]
    [TestCase("C:\\?.xlsx", 0, "C:\\\u5831\u544a.xlsx", "C:\\\u5831\u544a.xlsx")]
    public void FileMonikerPreservesAnsiAndCompleteUnicodePaths(string ansiPath, int parents, string? unicodePath, string expected)
    {
        byte[] bytes = CreateWorkbook(writer => Hyperlink(writer, 0, FileMoniker, data =>
            WriteFileMoniker(data, ansiPath, (ushort)parents, unicodePath, Encoding.GetEncoding(1252))));
        using var reader = Open(bytes);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetHyperlink(0)?.Address, Is.EqualTo(expected));
    }

    [Test]
    public void FileMonikerUsesExplicitLegacyCodePage()
    {
        var encoding = Encoding.GetEncoding(1251);
        const string path = "\u043e\u0442\u0447\u0451\u0442.xls";
        byte[] bytes = CreateWorkbook(
            writer => Hyperlink(writer, 0, FileMoniker, data =>
                WriteFileMoniker(data, path, 1, null, encoding)),
            1251);
        using var reader = Open(bytes);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetHyperlink(0)?.Address, Is.EqualTo("..\\" + path));
    }

    [TestCase("00000309-0000-0000-C000-000000000046")]
    [TestCase("00000304-0000-0000-C000-000000000046")]
    [TestCase("00000305-0000-0000-C000-000000000046")]
    [TestCase("01234567-89AB-CDEF-0123-456789ABCDEF")]
    public void UnsupportedMonikersDoNotHideEarlierLinksOrBlockLaterRecords(string classId)
    {
        byte[] bytes = CreateWorkbook(writer =>
        {
            StringHyperlink(writer, 0, "earlier");
            Hyperlink(writer, 0, classId, data => data.Write(new byte[] { 1, 2, 3 }), "ignored");
            StringHyperlink(writer, 1, "later");
        });
        using var reader = Open(bytes);
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetHyperlink(0)?.Address, Is.EqualTo("earlier"));
        Assert.That(reader.GetHyperlink(1)?.Address, Is.EqualTo("later"));
    }

    [Test]
    public void TooltipsMatchRangesRegardlessOfRecordOrder([Values] bool singlePass)
    {
        byte[] bytes = CreateWorkbook(writer =>
        {
            Tooltip(writer, 1, "second tip");
            StringHyperlink(writer, 0, "first");
            StringHyperlink(writer, 1, "second");
            Tooltip(writer, 0, "first tip");
        });
        using var source = new MemoryStream(bytes);
        using var reader = ExcelReaderFactory.CreateReader(source, new ExcelReaderConfiguration { SinglePassMode = singlePass });
        Assert.That(reader.Read(), Is.True);
        long position = source.Position;
        Assert.That(reader.GetHyperlink(0)?.Tooltip, Is.EqualTo("first tip"));
        Assert.That(reader.GetHyperlink(1)?.Tooltip, Is.EqualTo("second tip"));
        Assert.That(source.Position, Is.EqualTo(position));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(1), Is.EqualTo(11));
    }

    [Test]
    public void InvalidUrlLengthIsDeferredAndFailsExplicitly()
    {
        byte[] bytes = CreateWorkbook(writer => Hyperlink(writer, 0, UrlMoniker, data => data.Write(uint.MaxValue)));
        using var reader = Open(bytes);
        Assert.That(reader.Read(), Is.True);
        Assert.Throws<ExcelReaderException>(() => reader.GetHyperlink(0));
        Assert.That(reader.Read(), Is.True);
        Assert.That(reader.GetDouble(0), Is.EqualTo(10));
    }

    private static IExcelDataReader Open(byte[] bytes, bool singlePass = false) =>
        ExcelReaderFactory.CreateReader(new MemoryStream(bytes), new ExcelReaderConfiguration { SinglePassMode = singlePass });

    private static byte[] CreateWorkbook(Action<BinaryWriter> metadata, ushort? codePage = null)
    {
        using var output = new MemoryStream();
        using var writer = new BinaryWriter(output, Encoding.Unicode, leaveOpen: true);
        Record(writer, 0x809, data =>
        {
            data.Write((ushort)0x600);
            data.Write((ushort)0x10);
            data.Write(new byte[12]);
        });
        if (codePage != null)
            Record(writer, 0x42, data => data.Write(codePage.Value));
        for (int row = 0; row < 2; row++)
        {
            for (int column = 0; column < 2; column++)
            {
                Record(writer, 0x203, data =>
                {
                    data.Write((ushort)row);
                    data.Write((ushort)column);
                    data.Write((ushort)0);
                    data.Write((double)((row * 10) + column));
                });
            }
        }

        metadata(writer);
        Record(writer, 0xA, _ => { });
        return output.ToArray();
    }

    private static void Hyperlink(BinaryWriter writer, int column, string moniker, Action<BinaryWriter> body, string? location = null)
    {
        Record(writer, 0x1B8, data =>
        {
            Range(data, column);
            data.Write(new Guid("79EAC9D0-BAF9-11CE-8C82-00AA004BA90B").ToByteArray());
            data.Write(2);
            data.Write(location == null ? 1 : 9);
            data.Write(new Guid(moniker).ToByteArray());
            body(data);
            if (location != null)
                WideString(data, location);
        });
    }

    private static void StringHyperlink(BinaryWriter writer, int column, string address)
    {
        Record(writer, 0x1B8, data =>
        {
            Range(data, column);
            data.Write(new Guid("79EAC9D0-BAF9-11CE-8C82-00AA004BA90B").ToByteArray());
            data.Write(2);
            data.Write(0x101);
            WideString(data, address);
        });
    }

    private static void WriteFileMoniker(BinaryWriter writer, string ansiPath, ushort parents, string? unicodePath, Encoding encoding)
    {
        byte[] ansi = encoding.GetBytes(ansiPath + "\0");
        writer.Write(parents);
        writer.Write(ansi.Length);
        writer.Write(ansi);
        writer.Write(ushort.MaxValue);
        writer.Write((ushort)0xDEAD);
        writer.Write(new byte[20]);
        if (unicodePath == null)
        {
            writer.Write(0);
            return;
        }

        byte[] unicode = Encoding.Unicode.GetBytes(unicodePath);
        writer.Write(unicode.Length + 6);
        writer.Write(unicode.Length);
        writer.Write((ushort)3);
        writer.Write(unicode);
    }

    private static void Tooltip(BinaryWriter writer, int column, string text)
    {
        Record(writer, 0x800, data =>
        {
            data.Write((ushort)0x800);
            Range(data, column);
            data.Write(Encoding.Unicode.GetBytes(text + "\0"));
        });
    }

    private static void Range(BinaryWriter writer, int column)
    {
        writer.Write((ushort)0);
        writer.Write((ushort)0);
        writer.Write((ushort)column);
        writer.Write((ushort)column);
    }

    private static void WideString(BinaryWriter writer, string text)
    {
        writer.Write(text.Length + 1);
        writer.Write(Encoding.Unicode.GetBytes(text + "\0"));
    }

    private static void Record(BinaryWriter writer, ushort id, Action<BinaryWriter> write)
    {
        using var payload = new MemoryStream();
        using (var data = new BinaryWriter(payload, Encoding.Unicode, leaveOpen: true))
            write(data);
        writer.Write(id);
        writer.Write((ushort)payload.Length);
        writer.Write(payload.ToArray());
    }
}