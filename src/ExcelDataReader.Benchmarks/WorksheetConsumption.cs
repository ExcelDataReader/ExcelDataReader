using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class WorksheetConsumption
{
    private byte[] _data = [];
    private ExcelReaderConfiguration _configuration = null!;

    [Params("xls", "xlsx", "xlsb", "xml", "csv")]
    public string Format { get; set; } = string.Empty;

    [Params(false, true)]
    public bool SinglePass { get; set; }

    [Params(false, true)]
    public bool ReadValues { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        _configuration = new ExcelReaderConfiguration { SinglePassMode = SinglePass };
        if (Format is "xls" or "xlsx" or "xlsb")
        {
            using var stream = typeof(WorksheetConsumption).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks.10x10000." + Format)!;
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            _data = buffer.ToArray();
            return;
        }

        var text = new StringBuilder();
        if (Format == "xml")
            text.Append("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"><Worksheet ss:Name=\"Sheet1\"><Table>");
        for (int row = 0; row < 10000; row++)
        {
            if (Format == "csv")
            {
                text.Append("alpha,beta,gamma,delta,epsilon,zeta,eta,theta,iota,kappa\n");
            }
            else
            {
                text.Append("<Row>");
                for (int column = 0; column < 10; column++)
                    text.Append("<Cell><Data ss:Type=\"String\">alpha</Data></Cell>");
                text.Append("</Row>");
            }
        }

        if (Format == "xml")
            text.Append("</Table></Worksheet></Workbook>");
        _data = Encoding.UTF8.GetBytes(text.ToString());
    }

    [Benchmark]
    public long Read()
    {
        using var stream = new MemoryStream(_data, writable: false);
        using var reader = Format == "csv"
            ? ExcelReaderFactory.CreateCsvReader(stream, _configuration)
            : ExcelReaderFactory.CreateReader(stream, _configuration);
        long checksum = 0;
        while (reader.Read())
        {
            checksum++;
            if (!ReadValues)
                continue;
            for (int column = 0; column < reader.FieldCount; column++)
            {
                checksum += reader.GetValue(column) switch
                {
                    string value => value.Length,
                    double value => (long)value,
                    int value => value,
                    _ => 0,
                };
            }
        }

        return checksum;
    }
}