using System.Globalization;
using System.IO.Compression;
using System.Text;
using BenchmarkDotNet.Attributes;
using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class RepeatedSharedStrings
{
    private byte[] _data = [];
    private string _expected = string.Empty;

    [Params("xlsx", "xlsb")]
    public string Format { get; set; } = string.Empty;

    [Params("Ordinary", "Escaped", "Invalid")]
    public string Content { get; set; } = string.Empty;

    [Params("Repeated", "Unique")]
    public string Pattern { get; set; } = string.Empty;

    [Params(SharedStringStorageMode.Default, SharedStringStorageMode.SpillToDisk)]
    public SharedStringStorageMode Storage { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        string text = Content switch
        {
            "Escaped" => "first_x000A_second_x005F_x0041_",
            "Invalid" => "first_x00af_second_xG000_",
            _ => "first second ordinary string",
        };
        _expected = Content == "Escaped" ? "first\nsecond_x0041_" : text;
        if (Format == "xlsx")
        {
            using var buffer = new MemoryStream();
            var items = Enumerable.Repeat("<t>" + text + "</t>", Pattern == "Unique" ? 20000 : 1).ToList();
            if (Storage == SharedStringStorageMode.SpillToDisk)
                items.Add("<t>" + new string('p', 600000) + "</t>");
            byte[] template = AllocationTestWorkbook.CreateXlsx(items.ToArray(), true);
            buffer.Write(template, 0, template.Length);
            using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
            {
                archive.GetEntry("xl/worksheets/sheet1.xml")!.Delete();
                using var writer = new StreamWriter(archive.CreateEntry("xl/worksheets/sheet1.xml").Open());
                writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
                for (int row = 1; row <= 20000; row++)
                {
                    string index = row.ToString(CultureInfo.InvariantCulture);
                    string reference = (Pattern == "Unique" ? row - 1 : 0).ToString(CultureInfo.InvariantCulture);
                    writer.Write("<row r=\"" + index + "\"><c r=\"A" + index + "\" t=\"s\"><v>" + reference + "</v></c></row>");
                }

                writer.Write("</sheetData></worksheet>");
            }

            _data = buffer.ToArray();
        }
        else
        {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
            try
            {
                SharedStringWorkbook.Create(path, "xlsb", 20000, 64);
                using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
                {
                    archive.GetEntry("xl/sharedStrings.bin")!.Delete();
                    using (var writer = new BinaryWriter(archive.CreateEntry("xl/sharedStrings.bin").Open()))
                    {
                        for (int i = 0; i < (Pattern == "Unique" ? 20000 : 1); i++)
                        {
                            byte[] bytes = Encoding.Unicode.GetBytes(text);
                            writer.Write((byte)0x13);
                            writer.Write((byte)(5 + bytes.Length));
                            writer.Write((byte)0);
                            writer.Write(text.Length);
                            writer.Write(bytes);
                        }

                        if (Storage == SharedStringStorageMode.SpillToDisk)
                        {
                            var padding = new byte[5 + 1200000];
                            BitConverter.GetBytes(600000).CopyTo(padding, 1);
                            WriteRecord(writer, 0x13, padding);
                        }
                    }

                    archive.GetEntry("xl/worksheets/sheet1.bin")!.Delete();
                    using var sheet = new BinaryWriter(archive.CreateEntry("xl/worksheets/sheet1.bin").Open());
                    WriteRecord(sheet, 0x81, []);
                    WriteRecord(sheet, 0x91, []);
                    for (int row = 0; row < 5000; row++)
                    {
                        var header = new byte[17];
                        BitConverter.GetBytes(row).CopyTo(header, 0);
                        WriteRecord(sheet, 0, header);
                        for (int column = 0; column < 4; column++)
                        {
                            var cell = new byte[12];
                            BitConverter.GetBytes(column).CopyTo(cell, 0);
                            BitConverter.GetBytes(Pattern == "Unique" ? row * 4 + column : 0).CopyTo(cell, 8);
                            WriteRecord(sheet, 7, cell);
                        }
                    }

                    WriteRecord(sheet, 0x92, []);
                    WriteRecord(sheet, 0x82, []);
                }

                _data = File.ReadAllBytes(path);
            }
            finally
            {
                File.Delete(path);
            }
        }

        if (Read() != 20000L * _expected.Length)
            throw new InvalidOperationException("Shared string output changed.");
    }

    [Benchmark]
    public long Read()
    {
        using var reader = ExcelReaderFactory.CreateReader(
            new MemoryStream(_data, writable: false),
            new ExcelReaderConfiguration { SharedStringStorageMode = Storage, SharedStringSpillThreshold = 1024 * 1024 });
        long length = 0;
        while (reader.Read())
        {
            for (int column = 0; column < reader.FieldCount; column++)
            {
                string value = reader.GetString(column);
                if (value != _expected)
                    throw new InvalidOperationException("Shared string output changed.");
                length += value.Length;
            }
        }

        return length;
    }

    private static void WriteRecord(BinaryWriter writer, int id, byte[] data)
    {
        if (id >= 128)
        {
            writer.Write((byte)((id & 127) | 128));
            writer.Write((byte)(id >> 7));
        }
        else
        {
            writer.Write((byte)id);
        }

        int length = data.Length;
        do
        {
            writer.Write((byte)((length & 127) | (length > 127 ? 128 : 0)));
            length >>= 7;
        }
        while (length != 0);
        writer.Write(data);
    }
}