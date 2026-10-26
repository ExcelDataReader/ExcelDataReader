using System.Data;
using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class DataSetConversion
{
    private byte[] _data = [];

    [Params(false, true)]
    public bool FillMergedCells { get; set; }

    [Params(false, true)]
    public bool UseColumnDataType { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var xml = new StringBuilder("""
            <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
            <Worksheet ss:Name="Data"><Table>
            """);
        for (int row = 0; row < 2000; row++)
        {
            xml.Append("<Row>");
            for (int column = 0; column < 10; column++)
            {
                xml.Append(column == 0 && FillMergedCells ? "<Cell ss:MergeAcross=\"4\">" : "<Cell>");
                xml.Append("<Data ss:Type=\"Number\">");
                xml.Append((row * 10 + column).ToString(CultureInfo.InvariantCulture));
                xml.Append("</Data></Cell>");
                if (column == 0 && FillMergedCells)
                    column = 4;
            }

            xml.Append("</Row>");
        }

        xml.Append("</Table></Worksheet></Workbook>");
        _data = Encoding.UTF8.GetBytes(xml.ToString());
    }

    [Benchmark]
    public DataSet Convert()
    {
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(_data, writable: false));
        return reader.AsDataSet(new ExcelDataSetConfiguration
        {
            UseColumnDataType = UseColumnDataType,
            ConfigureDataTable = _ => new ExcelDataTableConfiguration { FillMergedCellsValue = FillMergedCells },
        });
    }
}