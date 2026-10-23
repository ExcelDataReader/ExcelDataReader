using System.Globalization;
using System.IO.Compression;
using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class OpenXmlOrdering
{
    private byte[] _data = [];

    [Params("xlsx", "xlsb")]
    public string Format { get; set; } = string.Empty;

    [Params("Ordered", "Overlap", "CrossBlock", "Sparse", "SparseOrdered", "Fragments", "Reverse")]
    public string Order { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        int[] rows = [.. Enumerable.Range(0, 10000)];
        if (Order == "Overlap")
        {
            for (int i = 0; i < rows.Length; i += 100)
                (rows[i], rows[i + 1]) = (rows[i + 1], rows[i]);
        }
        else if (Order == "CrossBlock")
        {
            (rows[0], rows[64]) = (rows[64], rows[0]);
        }
        else if (Order is "Sparse" or "SparseOrdered")
        {
            rows = [.. rows.Select(row => row * 3)];
            if (Order == "Sparse")
                (rows[0], rows[64]) = (rows[64], rows[0]);
        }
        else if (Order == "Fragments")
        {
            rows = [.. rows.SelectMany(row => Enumerable.Repeat(row, 20))];
        }
        else if (Order == "Reverse")
        {
            Array.Reverse(rows);
        }

        using var source = typeof(OpenXmlOrdering).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.10x10000." + Format)!;
        using var buffer = new MemoryStream();
        source.CopyTo(buffer);
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            string suffix = Format == "xlsx" ? ".xml" : ".bin";
            var entry = archive.Entries.First(item => item.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal) && item.FullName.EndsWith(suffix, StringComparison.Ordinal));
            string name = entry.FullName;
            entry.Delete();
            using var stream = archive.CreateEntry(name).Open();
            if (Format == "xlsx")
            {
                using var writer = new StreamWriter(stream, new UTF8Encoding(false));
                writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
                foreach (int row in rows)
                {
                    writer.Write("<row r=\"");
                    writer.Write((row + 1).ToString(CultureInfo.InvariantCulture));
                    writer.Write("\"><c><v>");
                    writer.Write((row + 0.25).ToString(CultureInfo.InvariantCulture));
                    writer.Write("</v></c></row>");
                }

                writer.Write("</sheetData></worksheet>");
            }
            else
            {
                using var writer = new BinaryWriter(stream);
                writer.Write(new byte[] { 0x91, 1, 0 });
                foreach (int row in rows)
                {
                    writer.Write((byte)0);
                    writer.Write((byte)12);
                    writer.Write(row);
                    writer.Write(0);
                    writer.Write((ushort)300);
                    writer.Write((byte)0);
                    writer.Write((byte)0b100000);
                    writer.Write((byte)5);
                    writer.Write((byte)16);
                    writer.Write(0);
                    writer.Write(0);
                    writer.Write(row + 0.25);
                }

                writer.Write(new byte[] { 0x92, 1, 0 });
            }
        }

        _data = buffer.ToArray();
        if (Read() != rows.Distinct().Sum(row => row + 0.25))
            throw new InvalidOperationException("The ordering benchmark did not return the expected cell values.");
    }

    [Benchmark]
    public int Open()
    {
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(_data, writable: false));
        return reader.RowCount;
    }

    [Benchmark]
    public double Read()
    {
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(_data, writable: false));
        double checksum = 0;
        while (reader.Read())
        {
            if (!reader.IsDBNull(0))
                checksum += reader.GetDouble(0);
        }

        return checksum;
    }
}