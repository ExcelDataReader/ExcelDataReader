using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class SpreadsheetValues
{
    private byte[] _data = [];
    private object _expected = null!;

    [Params("Number", "NumberDuration", "Date", "DateDuration", "TextDuration", "Boolean", "Text")]
    public string Values { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var (type, text, expected) = Values switch
        {
            "Number" => ("Number", "123.5", (object)123.5D),
            "NumberDuration" => ("Number", "2.5", TimeSpan.FromDays(2.5)),
            "Date" => ("DateTime", "2026-01-02T03:04:05", new DateTime(2026, 1, 2, 3, 4, 5)),
            "DateDuration" => ("DateTime", "1900-03-01T12:00:00", TimeSpan.FromDays(61.5)),
            "TextDuration" => ("String", "02:30:00", TimeSpan.FromHours(2.5)),
            "Boolean" => ("Boolean", "1", true),
            "Text" => ("String", "alpha", "alpha"),
            _ => throw new InvalidOperationException("Unknown value kind."),
        };
        _expected = expected;
        bool duration = Values.EndsWith("Duration", StringComparison.Ordinal);
        var xml = new StringBuilder("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"><Styles><Style ss:ID=\"value\"><NumberFormat ss:Format=\"");
        xml.Append(duration ? "[h]:mm:ss" : "General");
        xml.Append("\"/></Style></Styles><Worksheet ss:Name=\"Sheet1\"><Table>");
        for (int row = 0; row < 10000; row++)
        {
            xml.Append("<Row>");
            for (int column = 0; column < 10; column++)
                xml.Append("<Cell ss:StyleID=\"value\"><Data ss:Type=\"").Append(type).Append("\">").Append(text).Append("</Data></Cell>");
            xml.Append("</Row>");
        }

        xml.Append("</Table></Worksheet></Workbook>");
        _data = Encoding.UTF8.GetBytes(xml.ToString());
        if (Read() != 100000)
            throw new InvalidOperationException("Unexpected cell count.");
    }

    [Benchmark]
    public int Read()
    {
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(_data, writable: false));
        int count = 0;
        while (reader.Read())
        {
            for (int column = 0; column < reader.FieldCount; column++)
            {
                if (!Equals(reader.GetValue(column), _expected))
                    throw new InvalidOperationException("Unexpected cell value.");
                count++;
            }
        }

        return count;
    }
}