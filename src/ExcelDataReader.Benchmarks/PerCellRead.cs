using System.Globalization;
using System.IO.Compression;
using System.Text;
using BenchmarkDotNet.Attributes;
using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class PerCellRead
{
    private byte[] _data = [];

    [Params("Number", "Fraction", "Boolean", "Date", "Duration", "Text")]
    public string Values { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var buffer = new MemoryStream();
        var template = AllocationTestWorkbook.CreateXlsx([], false);
        buffer.Write(template, 0, template.Length);
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            archive.GetEntry("xl/_rels/workbook.xml.rels")!.Delete();
            using (var relationships = new StreamWriter(archive.CreateEntry("xl/_rels/workbook.xml.rels").Open()))
                relationships.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"sheet1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"styles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>");
            archive.GetEntry("xl/worksheets/sheet1.xml")!.Delete();
            using (var writer = new StreamWriter(archive.CreateEntry("xl/worksheets/sheet1.xml").Open()))
            {
                writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
                for (int row = 1; row <= 10000; row++)
                {
                    string index = row.ToString(CultureInfo.InvariantCulture);
                    writer.Write("<row r=\"" + index + "\">");
                    for (int column = 0; column < 10; column++)
                    {
                        string reference = ((char)('A' + column)).ToString() + index;
                        string attributes = Values switch
                        {
                            "Boolean" => " t=\"b\"",
                            "Date" => " s=\"1\"",
                            "Duration" => " s=\"2\"",
                            "Text" => " t=\"str\"",
                            _ => string.Empty,
                        };
                        string value = Values switch
                        {
                            "Boolean" => (row + column) % 2 == 0 ? "1" : "0",
                            "Text" => "alpha",
                            "Fraction" => index + ".5",
                            _ => index,
                        };
                        writer.Write("<c r=\"" + reference + "\"" + attributes + "><v>" + value + "</v></c>");
                    }

                    writer.Write("</row>");
                }

                writer.Write("</sheetData></worksheet>");
            }

            using var styles = new StreamWriter(archive.CreateEntry("xl/styles.xml").Open());
            styles.Write("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cellXfs count=\"3\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/><xf numFmtId=\"46\"/></cellXfs></styleSheet>");
        }

        _data = buffer.ToArray();
        long expected = Values switch
        {
            "Boolean" => 50000,
            "Text" => 500000,
            "Date" => 100000,
            "Duration" => 500050000,
            _ => 500050000,
        };
        if (Read() != expected)
            throw new InvalidOperationException("Unexpected cell values.");
        if (ReadTyped() != expected || ReadOnly() != 10000)
            throw new InvalidOperationException("Unexpected typed cell values or row count.");
    }

    [Benchmark]
    public long Read()
    {
        using var stream = new MemoryStream(_data, writable: false);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        long checksum = 0;
        while (reader.Read())
        {
            for (int column = 0; column < reader.FieldCount; column++)
            {
                checksum += reader.GetValue(column) switch
                {
                    double value when Values is "Number" or "Fraction" => (long)value,
                    bool value when Values == "Boolean" => value ? 1 : 0,
                    DateTime when Values == "Date" => 1,
                    TimeSpan value when Values == "Duration" => (long)value.TotalDays,
                    string value when Values == "Text" => value.Length,
                    _ => throw new InvalidOperationException("Unexpected cell type."),
                };
            }
        }

        return checksum;
    }

    [Benchmark]
    public int ReadOnly()
    {
        using var stream = new MemoryStream(_data, writable: false);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        int rows = 0;
        while (reader.Read())
            rows++;
        return rows;
    }

    [Benchmark]
    public long ReadTyped()
    {
        using var stream = new MemoryStream(_data, writable: false);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        long checksum = 0;
        while (reader.Read())
        {
            for (int column = 0; column < reader.FieldCount; column++)
            {
                checksum += Values switch
                {
                    "Number" or "Fraction" => (long)reader.GetDouble(column),
                    "Boolean" => reader.GetBoolean(column) ? 1 : 0,
                    "Date" => reader.GetDateTime(column).Year > 0 ? 1 : 0,
                    "Duration" => (long)((TimeSpan)reader.GetValue(column)).TotalDays,
                    "Text" => reader.GetString(column).Length,
                    _ => throw new InvalidOperationException("Unexpected cell type."),
                };
            }
        }

        return checksum;
    }

    [Benchmark]
    public long ReadRepeatedValues()
    {
        using var stream = new MemoryStream(_data, writable: false);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        long checksum = 0;
        while (reader.Read())
        {
            for (int column = 0; column < reader.FieldCount; column++)
            {
                var first = reader.GetValue(column);
                var second = reader.GetValue(column);
                if (!Equals(first, second))
                    throw new InvalidOperationException("Cell value changed between accesses.");
                checksum++;
            }
        }

        return checksum;
    }

    [Benchmark]
    public int ReadStyles()
    {
        using var stream = new MemoryStream(_data, writable: false);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        int checksum = 0;
        while (reader.Read())
        {
            for (int column = 0; column < reader.FieldCount; column++)
                checksum += reader.GetCellStyle(column).NumberFormatIndex;
        }

        return checksum;
    }
}