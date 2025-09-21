using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Plain string without backing storage. Used internally.
/// </summary>
internal sealed class XlsInternalString(string value) : IXlsString
{
    private readonly string stringValue = value;

    public string GetValue(Encoding encoding)
    {
        return stringValue;
    }
}
