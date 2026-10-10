using ExcelDataReader.Core.OpenXmlFormat.Records;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;

/// <summary>
/// Loads XLSB shared strings while each record buffer is owned by the reader.
/// </summary>
internal sealed class BiffSharedStringsReader(Stream stream) : BiffReader(stream)
{
    private const int StringItem = 0x13;

    private ISharedStringStore? _store;

    public void Load(ISharedStringStore store)
    {
        _store = store ?? throw new ArgumentNullException(nameof(store));
        try
        {
            while (base.Read() != null)
            {
            }
        }
        finally
        {
            _store = null;
        }
    }

    public override Record? Read() =>
        throw new InvalidOperationException("XLSB shared strings are loaded synchronously.");

    protected override Record ReadOverride(byte[] buffer, uint recordId, uint recordLength)
    {
        switch (recordId)
        {
            case StringItem:
                // Flags byte, character count and UTF-16 text; trailing rich text and phonetic data are ignored.
                if (recordLength < 1 + 4)
                    throw new ExcelReaderException(Errors.ErrorBiffStringSize);
                uint length = GetDWord(buffer, 1);
                if (length > (recordLength - (1 + 4)) / 2)
                    throw new ExcelReaderException(Errors.ErrorBiffStringSize);
                (_store ?? throw new InvalidOperationException("Shared strings are not being loaded."))
                    .AddUtf16(buffer, 1 + 4, (int)length);
                break;
        }

        return Record.Default;
    }
}
