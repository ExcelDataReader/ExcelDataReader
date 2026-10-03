using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks;

[MemoryDiagnoser]
public class ReadEncryptedFiles
{
    private byte[] _data = [];
    private ExcelReaderConfiguration _configuration;

    [Params(
        "agile_AES128_SHA1_CBC_pwd_password.xlsx",
        "standard_AES128_SHA1_ECB_pwd_password.xlsx",
        "Issue242_StdRc4PwdPassword.xls",
        "Issue242_XorPwdPassword.xls")]
    public string File { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        using var resource = typeof(ReadEncryptedFiles).Assembly.GetManifestResourceStream("ExcelDataReader.Benchmarks." + File)
            ?? throw new InvalidOperationException("Missing embedded resource " + File);
        using var buffer = new MemoryStream();
        resource.CopyTo(buffer);
        _data = buffer.ToArray();
        _configuration = new ExcelReaderConfiguration { Password = "password" };
    }

    [Benchmark]
    public void OpenReader()
    {
        using var reader = ExcelReaderFactory.CreateReader(new MemoryStream(_data, writable: false), _configuration);
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
