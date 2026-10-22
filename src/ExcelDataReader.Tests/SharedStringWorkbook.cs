using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace ExcelDataReader.TestFixtures;

internal static class SharedStringWorkbook
{
    public static string Value(int index, int length, bool unicode = false)
    {
        var chars = new char[length];
        string prefix = index.ToString("D10", CultureInfo.InvariantCulture);
        prefix.CopyTo(0, chars, 0, Math.Min(prefix.Length, length));
        uint state = (uint)index + 741;
        for (int i = prefix.Length; i < length; i++)
        {
            state = unchecked(state * 1664525 + 1013904223);
            chars[i] = unicode ? (char)(0x400 + state % 64) : (char)('a' + state % 26);
        }

        return new string(chars);
    }

    public static int Reference(int position, int count, string pattern) => pattern switch
    {
        "Sequential" => position,
        "Permuted" => (int)(((long)position * (count - 1) + 741) % count),
        "Random" => RandomReference(position, count),
        "Hot" when position < count => position % Math.Min(count, 256),
        "Hot" => position - count,
        "Sparse" => position % Math.Min(count, 256),
        _ => throw new ArgumentException("Unknown reference pattern.", nameof(pattern)),
    };

    public static int References(int count, string pattern) => pattern switch
    {
        "Hot" => checked(count * 2),
        "Sparse" => Math.Min(count, 10000),
        "Sequential" or "Permuted" or "Random" => count,
        _ => throw new ArgumentException("Unknown reference pattern.", nameof(pattern)),
    };

    public static ulong Hash(ulong hash, string value)
    {
        foreach (char c in value)
            hash = unchecked((hash ^ c) * 1099511628211);
        return hash;
    }

    public static void Create(string path, string format, int count, int length, string pattern = "Sequential", bool unicode = false)
    {
        if (count < 1 || length < 10 || length > 4096)
            throw new ArgumentOutOfRangeException(nameof(count), "Positive count and length between 10 and 4096 required.");
        int references = References(count, pattern);
        using var file = new FileStream(path, FileMode.Create, FileAccess.ReadWrite);
        if (format == "xls")
        {
            WriteXls(file, count, length, references, pattern, unicode);
            return;
        }

        if (format is not ("xlsx" or "xlsb"))
            throw new ArgumentException("Unknown workbook format.", nameof(format));
        using var archive = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: true);
        string suffix = format == "xlsx" ? "xml" : "bin";
        using (var writer = XmlWriter.Create(archive.CreateEntry("xl/_rels/workbook." + suffix + ".rels").Open(), new XmlWriterSettings { CloseOutput = true }))
        {
            writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
            Relationship(writer, "sheet1", "worksheet", "worksheets/sheet1." + suffix);
            Relationship(writer, "sst", "sharedStrings", "sharedStrings." + suffix);
            writer.WriteEndElement();
        }

        if (format == "xlsx")
            WriteXlsx(archive, count, length, references, pattern, unicode);
        else
            WriteXlsb(archive, count, length, references, pattern, unicode);
    }

    private static int RandomReference(int position, int count)
    {
        int step = 104729;
        while (GreatestCommonDivisor(step, count) != 1)
            step += 2;
        return (int)(((long)position * step + 741) % count);
    }

    private static int GreatestCommonDivisor(int a, int b)
    {
        while (b != 0)
        {
            int next = a % b;
            a = b;
            b = next;
        }

        return a;
    }

    private static void Relationship(XmlWriter writer, string id, string type, string target)
    {
        writer.WriteStartElement("Relationship");
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + type);
        writer.WriteAttributeString("Target", target);
        writer.WriteEndElement();
    }

    private static void WriteXlsx(ZipArchive archive, int count, int length, int references, string pattern, bool unicode)
    {
        const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using (var writer = XmlWriter.Create(archive.CreateEntry("xl/workbook.xml").Open(), new XmlWriterSettings { CloseOutput = true }))
        {
            writer.WriteStartElement("workbook", ns);
            writer.WriteStartElement("sheets", ns);
            writer.WriteStartElement("sheet", ns);
            writer.WriteAttributeString("name", "Sheet1");
            writer.WriteAttributeString("sheetId", "1");
            writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "sheet1");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        using (var writer = XmlWriter.Create(archive.CreateEntry("xl/sharedStrings.xml").Open(), new XmlWriterSettings { CloseOutput = true }))
        {
            writer.WriteStartElement("sst", ns);
            writer.WriteAttributeString("uniqueCount", count.ToString(CultureInfo.InvariantCulture));
            for (int i = 0; i < count; i++)
            {
                writer.WriteStartElement("si", ns);
                writer.WriteElementString("t", ns, Value(i, length, unicode));
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        using (var writer = XmlWriter.Create(archive.CreateEntry("xl/worksheets/sheet1.xml").Open(), new XmlWriterSettings { CloseOutput = true }))
        {
            writer.WriteStartElement("worksheet", ns);
            writer.WriteStartElement("sheetData", ns);
            for (int i = 0; i < references; i += 4)
            {
                writer.WriteStartElement("row", ns);
                string row = (i / 4 + 1).ToString(CultureInfo.InvariantCulture);
                writer.WriteAttributeString("r", row);
                for (int c = 0; c < 4 && i + c < references; c++)
                {
                    writer.WriteStartElement("c", ns);
                    writer.WriteAttributeString("r", ((char)('A' + c)).ToString() + row);
                    writer.WriteAttributeString("t", "s");
                    writer.WriteElementString("v", ns, Reference(i + c, count, pattern).ToString(CultureInfo.InvariantCulture));
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteEndElement();
        }
    }

    private static void WriteXlsb(ZipArchive archive, int count, int length, int references, string pattern, bool unicode)
    {
        using (var writer = new BinaryWriter(archive.CreateEntry("xl/workbook.bin").Open(), Encoding.Unicode))
        {
            BinaryRecord(writer, 0x83, []);
            using var buffer = new MemoryStream();
            using var data = new BinaryWriter(buffer, Encoding.Unicode, leaveOpen: true);
            data.Write(0);
            data.Write(1);
            WideString(data, "sheet1");
            WideString(data, "Sheet1");
            BinaryRecord(writer, 0x9C, buffer.ToArray());
            BinaryRecord(writer, 0x84, []);
        }

        using (var writer = new BinaryWriter(archive.CreateEntry("xl/sharedStrings.bin").Open(), Encoding.Unicode))
        {
            var header = new byte[8];
            BitConverter.GetBytes(references).CopyTo(header, 0);
            BitConverter.GetBytes(count).CopyTo(header, 4);
            BinaryRecord(writer, 0x9F, header);
            for (int i = 0; i < count; i++)
            {
                byte[] bytes = new byte[5 + length * 2];
                BitConverter.GetBytes(length).CopyTo(bytes, 1);
                Encoding.Unicode.GetBytes(Value(i, length, unicode), 0, length, bytes, 5);
                BinaryRecord(writer, 0x13, bytes);
            }

            BinaryRecord(writer, 0xA0, []);
        }

        using (var writer = new BinaryWriter(archive.CreateEntry("xl/worksheets/sheet1.bin").Open(), Encoding.Unicode))
        {
            BinaryRecord(writer, 0x81, []);
            BinaryRecord(writer, 0x91, []);
            for (int i = 0; i < references; i += 4)
            {
                var row = new byte[17];
                BitConverter.GetBytes(i / 4).CopyTo(row, 0);
                BinaryRecord(writer, 0, row);
                for (int c = 0; c < 4 && i + c < references; c++)
                {
                    var cell = new byte[12];
                    BitConverter.GetBytes(c).CopyTo(cell, 0);
                    BitConverter.GetBytes(Reference(i + c, count, pattern)).CopyTo(cell, 8);
                    BinaryRecord(writer, 7, cell);
                }
            }

            BinaryRecord(writer, 0x92, []);
            BinaryRecord(writer, 0x82, []);
        }
    }

    private static void WideString(BinaryWriter writer, string value)
    {
        writer.Write(value.Length);
        writer.Write(Encoding.Unicode.GetBytes(value));
    }

    private static void BinaryRecord(BinaryWriter writer, uint id, byte[] data)
    {
        Variable(writer, id);
        Variable(writer, (uint)data.Length);
        writer.Write(data);
    }

    private static void Variable(BinaryWriter writer, uint value)
    {
        do
        {
            byte b = (byte)(value & 127);
            value >>= 7;
            writer.Write(value == 0 ? b : (byte)(b | 128));
        }
        while (value != 0);
    }

    private static void WriteXls(Stream file, int count, int length, int references, string pattern, bool unicode)
    {
        using var writer = new BinaryWriter(file, Encoding.Unicode, leaveOpen: true);
        Bof(writer, 5);
        int sheets = (references + 262143) / 262144;
        var offsets = new long[sheets];
        for (int s = 0; s < sheets; s++)
        {
            string name = "Sheet" + (s + 1).ToString(CultureInfo.InvariantCulture);
            var bound = new byte[8 + name.Length];
            bound[6] = (byte)name.Length;
            Encoding.ASCII.GetBytes(name).CopyTo(bound, 8);
            offsets[s] = file.Position + 4;
            XlsRecord(writer, 0x85, bound);
        }

        using (var buffer = new MemoryStream())
        using (var data = new BinaryWriter(buffer, Encoding.Unicode, leaveOpen: true))
        {
            data.Write(references);
            data.Write(count);
            ushort id = 0xFC;
            for (int i = 0; i < count; i++)
            {
                int size = 3 + length * (unicode ? 2 : 1);
                if (buffer.Length + size > 8224)
                {
                    XlsRecord(writer, id, buffer.ToArray());
                    id = 0x3C;
                    buffer.SetLength(0);
                    buffer.Position = 0;
                }

                data.Write((ushort)length);
                data.Write((byte)(unicode ? 1 : 0));
                data.Write((unicode ? Encoding.Unicode : Encoding.ASCII).GetBytes(Value(i, length, unicode)));
            }

            XlsRecord(writer, id, buffer.ToArray());
        }

        XlsRecord(writer, 0xA, []);
        for (int s = 0; s < sheets; s++)
        {
            long start = file.Position;
            file.Position = offsets[s];
            writer.Write((uint)start);
            file.Position = start;
            Bof(writer, 0x10);
            int end = Math.Min(references, (s + 1) * 262144);
            for (int i = s * 262144; i < end; i++)
            {
                var cell = new byte[10];
                BitConverter.GetBytes((ushort)((i % 262144) / 4)).CopyTo(cell, 0);
                BitConverter.GetBytes((ushort)(i % 4)).CopyTo(cell, 2);
                BitConverter.GetBytes(Reference(i, count, pattern)).CopyTo(cell, 6);
                XlsRecord(writer, 0xFD, cell);
            }

            XlsRecord(writer, 0xA, []);
        }
    }

    private static void Bof(BinaryWriter writer, ushort type)
    {
        var data = new byte[16];
        BitConverter.GetBytes((ushort)0x600).CopyTo(data, 0);
        BitConverter.GetBytes(type).CopyTo(data, 2);
        XlsRecord(writer, 0x809, data);
    }

    private static void XlsRecord(BinaryWriter writer, ushort id, byte[] data)
    {
        writer.Write(id);
        writer.Write((ushort)data.Length);
        writer.Write(data);
    }
}
