using System.Text;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// The font with index 4 is omitted in all BIFF versions. This means the first four fonts have zero-based indexes, and the fifth font and all following fonts are referenced with one-based indexes.
    /// </summary>
    internal class XlsBiffFont : XlsBiffRecord
    {
        private readonly IXlsString _fontName;

        internal XlsBiffFont(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            if (Id == BIFFRECORDTYPE.FONT_V34)
            {
                _fontName = new XlsShortByteString(bytes, ContentOffset + 6);
            }
            else if (Id == BIFFRECORDTYPE.FONT && biffVersion == 2)
            {
                _fontName = new XlsShortByteString(bytes, ContentOffset + 4);
            }
            else if (Id == BIFFRECORDTYPE.FONT && biffVersion == 5)
            {
                _fontName = new XlsShortByteString(bytes, ContentOffset + 14);
            }
            else if (Id == BIFFRECORDTYPE.FONT && biffVersion == 8)
            {
                _fontName = new XlsShortUnicodeString(bytes, ContentOffset + 14);
            }
            else
            {
                _fontName = new XlsInternalString(string.Empty);
            }

            if (Id == BIFFRECORDTYPE.FONT && biffVersion >= 5)
            {
                // Encodings were mapped by correlating this:
                // https://docs.microsoft.com/en-us/windows/desktop/intl/code-page-identifiers
                // with the FONT record character set table here:
                // https://www.openoffice.org/sc/excelfileformat.pdf
                var byteStringCharacterSet = ReadByte(12);
                switch (byteStringCharacterSet)
                {
                    case 0: // ANSI Latin
                    case 1: // System default
                        ByteStringEncoding = EncodingHelper.GetEncoding(1252);
                        break;
                    case 77: // Apple roman
                        ByteStringEncoding = EncodingHelper.GetEncoding(10000);
                        break;
                    case 128: // ANSI Japanese Shift-JIS
                        ByteStringEncoding = EncodingHelper.GetEncoding(932);
                        break;
                    case 129: // ANSI Korean (Hangul)
                        ByteStringEncoding = EncodingHelper.GetEncoding(949);
                        break;
                    case 130: // ANSI Korean (Johab)
                        ByteStringEncoding = EncodingHelper.GetEncoding(1361);
                        break;
                    case 134: // ANSI Chinese Simplified GBK
                        ByteStringEncoding = EncodingHelper.GetEncoding(936);
                        break;
                    case 136: // ANSI Chinese Traditional BIG5
                        ByteStringEncoding = EncodingHelper.GetEncoding(950);
                        break;
                    case 161: // ANSI Greek
                        ByteStringEncoding = EncodingHelper.GetEncoding(1253);
                        break;
                    case 162: // ANSI Turkish
                        ByteStringEncoding = EncodingHelper.GetEncoding(1254);
                        break;
                    case 163: // ANSI Vietnamese
                        ByteStringEncoding = EncodingHelper.GetEncoding(1258);
                        break;
                    case 177: // ANSI Hebrew
                        ByteStringEncoding = EncodingHelper.GetEncoding(1255);
                        break;
                    case 178: // ANSI Arabic
                        ByteStringEncoding = EncodingHelper.GetEncoding(1256);
                        break;
                    case 186: // ANSI Baltic
                        ByteStringEncoding = EncodingHelper.GetEncoding(1257);
                        break;
                    case 204: // ANSI Cyrillic
                        ByteStringEncoding = EncodingHelper.GetEncoding(1251);
                        break;
                    case 222: // ANSI Thai
                        ByteStringEncoding = EncodingHelper.GetEncoding(874);
                        break;
                    case 238: // ANSI Latin II
                        ByteStringEncoding = EncodingHelper.GetEncoding(1250);
                        break;
                    case 255: // OEM Latin
                        ByteStringEncoding = EncodingHelper.GetEncoding(850);
                        break;
                }
            }
        }

        public Encoding ByteStringEncoding { get; }

        public string GetFontName(Encoding encoding) => _fontName.GetValue(encoding);
    }
}
