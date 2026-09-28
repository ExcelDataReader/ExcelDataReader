using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

/// <summary>
/// Reads every value in all sheets of real-world test workbooks with mixed content
/// (numbers, dates, formulas, inline and shared strings, blanks).
/// </summary>
[MemoryDiagnoser]
public class ReadRealWorldFiles
{
    private byte[] _data = [];

    private ExcelReaderConfiguration _configuration;

    [Params(
        "BigFormatted.xls",
        "BigFormatted.xlsx",
        "Issue263.xls",
        "OldIssue11572_CodePage.xls",
        "Issue173.xls",
        "InvalidByteOrderValueInHeader.xls",
        "LotsOfSheets.xls",
        "LotsOfSheets.xlsx",
        "LotsOfSheets.xlsb",
        "agile_AES128_SHA1_CBC_pwd_password.xlsx",
        "standard_AES128_SHA1_ECB_pwd_password.xlsx",
        "Issue242_StdRc4PwdPassword.xls",
        "Issue242_XorPwdPassword.xls")]
    public string File { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        using var resource = typeof(ReadRealWorldFiles).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks." + File)
            ?? throw new InvalidOperationException("Missing embedded resource " + File);
        using var buffer = new MemoryStream();
        resource.CopyTo(buffer);
        _data = buffer.ToArray();

        // Encrypted test workbooks all use the password "password".
        _configuration = File.IndexOf("pwd", StringComparison.OrdinalIgnoreCase) >= 0
            ? new ExcelReaderConfiguration { Password = "password" }
            : null;
    }

    [Benchmark]
    public int ReadAllValues()
    {
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(_data, writable: false), _configuration);
        var nonNull = 0;
        do
        {
            while (reader.Read())
            {
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    if (reader.GetValue(i) != null)
                        nonNull++;
                }
            }
        }
        while (reader.NextResult());

        return nonNull;
    }
}
