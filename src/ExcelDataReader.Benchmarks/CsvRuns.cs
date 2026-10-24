using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class CsvRuns
{
    private byte[] _data = [];
    private ExcelReaderConfiguration _configuration = null!;

    [Params("short", "long", "quoted", "escaped", "trimmed")]
    public string Content { get; set; } = string.Empty;

    [Params(false, true)]
    public bool DetectSeparator { get; set; }

    [Params(",", "||")]
    public string Separator { get; set; } = ",";

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        string field = Content switch
        {
            "long" => new string('x', 1024),
            "quoted" => "\"" + new string('x', 1024) + ",\r\nend\"",
            "escaped" => "\"a\"\"b\"\"c\"",
            "trimmed" => "  alpha   ",
            _ => "alpha",
        };
        var text = new StringBuilder();
        for (int row = 0; row < 1000; row++)
            text.Append(field).Append(Separator).Append(field).Append(Separator).Append(field).Append("\r\n");
        _data = Encoding.UTF8.GetBytes(text.ToString());
        _configuration = new ExcelReaderConfiguration
        {
            AutodetectSeparators = DetectSeparator ? [",", ";", "\t", Separator == "," ? "|" : Separator, "#"] : [Separator],
        };

        int decodedLength = Content switch
        {
            "long" => 1024,
            "quoted" => 1030,
            _ => 5,
        };
        if (Read() != 1000L * (1 + 3 * decodedLength))
            throw new InvalidOperationException("CSV benchmark checksum mismatch.");
    }

    [Benchmark]
    public long Read()
    {
        using var stream = new MemoryStream(_data, false);
        using var reader = ExcelReaderFactory.CreateCsvReader(stream, _configuration);
        long checksum = 0;
        while (reader.Read())
        {
            checksum++;
            for (int column = 0; column < reader.FieldCount; column++)
                checksum += reader.GetString(column).Length;
        }

        return checksum;
    }
}