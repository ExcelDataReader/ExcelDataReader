using System.Text;

namespace ExcelDataReader.Core;

internal sealed class EncodingHelper
{
    public static string DecodeLatin1(byte[] bytes, int offset, int count)
    {
#if NET8_0_OR_GREATER
        return Encoding.Latin1.GetString(bytes.AsSpan(offset, count));
#elif NETSTANDARD2_1_OR_GREATER
        return string.Create(count, (Bytes: bytes, Offset: offset), static (destination, state) =>
        {
            for (int i = 0; i < destination.Length; i++)
                destination[i] = (char)state.Bytes[state.Offset + i];
        });
#else
        var characters = new char[count];
        for (int i = 0; i < count; i++)
            characters[i] = (char)bytes[offset + i];
        return new string(characters);
#endif
    }

    public static Encoding GetEncoding(ushort codePage)
    {
        int encodingCodePage = codePage switch
        {
            // BIFF2/3 use special IDs for Mac Roman and Windows-1252.
            32768 => 10000,
            32769 => 1252,
            _ => codePage
        };

        if (encodingCodePage == 0)
            throw new NotSupportedException("Code page 0 does not identify a specific encoding.");

#if NET10_0_OR_GREATER
        try
        {
            return Encoding.GetEncoding(encodingCodePage);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException)
        {
            var encoding = CodePagesEncodingProvider.Instance.GetEncoding(encodingCodePage);
            if (encoding is not null)
                return encoding;

            throw;
        }
#else
        return Encoding.GetEncoding(encodingCodePage);
#endif
    }
}
