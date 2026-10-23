using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;

/// <summary>
/// Loads XLSB shared strings while each record buffer is owned by the reader.
/// </summary>
internal sealed class BiffSharedStringsReader(Stream stream) : BiffReader(stream)
{
    private const int StringItem = 0x13;
    private char[]? _characters;

    public void Load(ISharedStringStore store)
    {
        _ = store ?? throw new ArgumentNullException(nameof(store));

        while (TryReadFrame(out var buffer, out var recordId, out var recordLength))
        {
            if (recordId != StringItem)
                continue;

            // Flags byte, character count and UTF-16 text; trailing rich text and phonetic data are ignored.
            if (recordLength < 1 + 4)
                throw new ExcelReaderException(Errors.ErrorBiffStringSize);
            uint length = GetDWord(buffer, 1);
            if (length > (recordLength - (1 + 4)) / 2)
                throw new ExcelReaderException(Errors.ErrorBiffStringSize);
            AddString(store, buffer, 1 + 4, (int)length);
        }
    }

    private void AddString(ISharedStringStore store, byte[] buffer, int offset, int length)
    {
        for (int i = 0; i + 1 < length; i++)
        {
            int position = offset + i * 2;
            if (buffer[position] != '_' || buffer[position + 1] != 0 ||
                buffer[position + 2] != 'x' || buffer[position + 3] != 0)
            {
                continue;
            }

            // Preserve raw UTF-16 code units, including unpaired surrogates, for disk storage.
            var characters = _characters;
            if (characters == null || characters.Length < length)
                _characters = characters = new char[length];
            for (int j = 0; j < length; j++)
                characters[j] = (char)(buffer[offset + j * 2] | buffer[offset + j * 2 + 1] << 8);
            string value = new(characters, 0, length);
            store.Add(Helpers.ConvertEscapeChars(value));
            return;
        }

        store.AddUtf16(buffer, offset, length);
    }
}