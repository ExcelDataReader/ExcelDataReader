using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat;

internal abstract class RecordReader : IDisposable
{
    ~RecordReader()
    {
        Dispose(false);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }       

    public abstract Record? Read();

    protected abstract void Dispose(bool disposing);
}
