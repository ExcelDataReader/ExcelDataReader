using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;

/// <summary>
/// Loads XLSB shared strings while each record buffer is owned by the reader.
/// </summary>
internal sealed class BiffSharedStringsReader(Stream stream) : BiffReader(stream)
{
    private const int StringItem = 0x13;

    public void Load(ISharedStringStore store)
    {
        _ = store ?? throw new ArgumentNullException(nameof(store));

        while (TryReadFrame(out var buffer, out var recordId, out var recordLength))
        {
            try
            {
                if (recordId != StringItem)
                    continue;

                // Flags byte, character count and UTF-16 text; trailing rich text and phonetic data are ignored.
                if (recordLength < 1 + 4)
                    throw new ExcelReaderException(Errors.ErrorBiffStringSize);
                uint length = GetDWord(buffer, 1);
                if (length > (recordLength - (1 + 4)) / 2)
                    throw new ExcelReaderException(Errors.ErrorBiffStringSize);
                store.AddUtf16(buffer, 1 + 4, (int)length);
            }
            finally
            {
                ReleaseFrame(buffer);
            }
        }
    }
}
