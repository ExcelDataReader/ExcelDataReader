using System.Text;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// [MS-XLS] 2.5.240 ShortXLUnicodeString
/// Byte-sized string, stored as single or multibyte unicode characters.
/// </summary>
internal sealed class XlsShortUnicodeString(byte[] bytes, uint offset) : IXlsString
{
    private readonly byte[] _bytes = bytes;
    private readonly uint _offset = offset;

    public ushort CharacterCount => _bytes[_offset];

    /// <summary>
    /// Gets a value indicating whether the string is a multibyte string or not.
    /// </summary>
    public bool IsMultiByte => (_bytes[_offset + 1] & 0x01) != 0;

    public string GetValue(Encoding encoding)
    {
        if (CharacterCount == 0)
        {
            return string.Empty;
        }

        int available = _bytes.Length - (int)_offset - 2;
        int needed = IsMultiByte ? CharacterCount * 2 : CharacterCount;
        if (available < needed)
        {
            throw new ExcelReaderException(Errors.ErrorBiffStringSize);
        }

        if (IsMultiByte)
        {
            return Encoding.Unicode.GetString(_bytes, (int)_offset + 2, CharacterCount * 2);
        }

        // In BIFF8 single-byte strings each byte value IS the Unicode code point,
        // so a direct (char)byte cast is the exact and encoding-free conversion.
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        return string.Create(CharacterCount, (Bytes: _bytes, Start: (int)_offset + 2), static (span, state) =>
        {
            for (int i = 0; i < span.Length; i++)
                span[i] = (char)state.Bytes[state.Start + i];
        });
#else
        var chars = new char[CharacterCount];
        int start = (int)_offset + 2;
        for (int i = 0; i < CharacterCount; i++)
            chars[i] = (char)_bytes[start + i];
        return new string(chars);
#endif
    }
}
