using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class SpreadsheetStyles
{
    private byte[] _data = [];

    [Params(1, 16, 256)]
    public int Styles { get; set; }

    [Params(1, 10000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var xml = new StringBuilder("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"><Styles>");
        for (int i = 0; i < Styles; i++)
            xml.Append("<Style ss:ID=\"s").Append(i).Append("\"><NumberFormat ss:Format=\"0.00\"/></Style>");
        xml.Append("</Styles><Worksheet ss:Name=\"Sheet1\"><Table>");
        for (int i = 0; i < Rows; i++)
            xml.Append("<Row><Cell ss:StyleID=\"s").Append(i % Styles).Append("\"><Data ss:Type=\"Number\">123.5</Data></Cell></Row>");
        xml.Append("</Table></Worksheet></Workbook>");
        _data = Encoding.UTF8.GetBytes(xml.ToString());
    }

    [Benchmark]
    public double Read()
    {
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(_data, false));
        double sum = 0;
        while (reader.Read())
            sum += reader.GetDouble(0);
        if (sum != Rows * 123.5)
            throw new InvalidOperationException("SpreadsheetML checksum mismatch.");
        return sum;
    }
}