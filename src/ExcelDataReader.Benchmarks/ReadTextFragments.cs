using System.Text;
using BenchmarkDotNet.Attributes;
using ExcelDataReader.TestFixtures;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class ReadTextFragments
{
    private byte[] _data = [];

    [Params("InlineSingle", "InlineRich", "SharedSingle", "SharedRich", "XmlSingle", "XmlFragments", "EscapeOrdinary", "EscapeValid", "EscapeInvalid", "EscapeLong")]
    public string Content { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        string fragment = Content switch
        {
            "InlineRich" or "SharedRich" => "<r><t>first</t></r><r><t xml:space=\"preserve\"> second</t></r>",
            "XmlSingle" => "<Data ss:Type=\"String\">first second</Data>",
            "XmlFragments" => "<Data ss:Type=\"String\">first<![CDATA[ second]]><B> third</B></Data>",
            "EscapeValid" => "<t>first_x000A_second_x0041__x005F_x0041_</t>",
            "EscapeInvalid" => "<t>first_x00af_second_xG000_third_x004</t>",
            "EscapeLong" => "<t>" + new string('a', 4096) + "_x000A_" + new string('b', 4096) + "</t>",
            _ => "<t>first second</t>",
        };
        var items = Enumerable.Repeat(fragment, 1000).ToArray();
        _data = Content.StartsWith("Xml", StringComparison.Ordinal)
            ? AllocationTestWorkbook.CreateSpreadsheetXml(items)
            : AllocationTestWorkbook.CreateXlsx(items, Content.StartsWith("Shared", StringComparison.Ordinal));
    }

    [Benchmark]
    public int ReadValues()
    {
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(_data, writable: false));
        int length = 0;
        while (reader.Read())
            length += reader.GetString(0)?.Length ?? 0;
        return length;
    }
}
