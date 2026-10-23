using System.Text;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat;

// Unicode is decoded immediately. Legacy character bytes are owned until a later
// CODEPAGE record supplies their encoding; no record header or borrowed data is retained.
internal sealed class XlsStringData
{
    private readonly string? _value;
    private readonly byte[]? _characters;
    private readonly int _characterCount;

    public XlsStringData(string value) => _value = value;

    private XlsStringData(byte[] characters, int characterCount)
    {
        _characters = characters;
        _characterCount = characterCount;
    }

    public static XlsStringData Read(XlsBiffFrame frame, int offset, bool shortCount, bool unicode)
    {
        int count = shortCount ? frame.ReadStringByteAt(offset) : frame.ReadUInt16At(offset);
        offset += shortCount ? 1 : 2;
        if (unicode)
        {
            if (count == 0)
                return new XlsStringData(string.Empty);

            bool multiByte = (frame.ReadStringByteAt(offset++) & 1) != 0;
            int byteCount = multiByte ? count * 2 : count;
            if (frame.RecordSize + 4 - offset < byteCount)
                throw new ExcelReaderException(Errors.ErrorBiffStringSize);

            if (multiByte)
                return new XlsStringData(Encoding.Unicode.GetString(frame.Bytes, offset, byteCount));

            var characters = new char[count];
            for (int i = 0; i < count; i++)
                characters[i] = (char)frame.Bytes[offset + i];
            return new XlsStringData(new string(characters));
        }

        int available = Math.Min(count * 2, frame.RecordSize + 4 - offset);
        var bytes = new byte[available];
        Buffer.BlockCopy(frame.Bytes, offset, bytes, 0, available);
        return new XlsStringData(bytes, count);
    }

    public string GetValue(Encoding encoding) =>
        _value ?? encoding.GetString(_characters!, 0, _characterCount * (Helpers.IsSingleByteEncoding(encoding) ? 1 : 2));
}