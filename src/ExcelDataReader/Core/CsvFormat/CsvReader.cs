using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelDataReader.Core.CsvFormat
{
    internal class CsvReader
    {
        private const int BufferSize = 1024;

        public CsvReader(Stream stream, char separator, Encoding fallbackEncoding)
        {
            BaseStream = stream;
            Separator = separator;
            Buffer = new byte[BufferSize];
            ReadBuffer();
            
            var encoding = GetEncodingFromBom(Buffer, out var bomLength);
            if (encoding != null)
            {
                Advance(bomLength);
                Encoding = encoding;
            }
            else
            {
                Encoding = fallbackEncoding;
            }

            Decoder = Encoding.GetDecoder();
            Decoder.Fallback = new DecoderExceptionFallback();

            MaxCharBytes = Encoding.GetMaxByteCount(1);
        }

        public Encoding Encoding { get; }

        private Decoder Decoder { get; }

        private Stream BaseStream { get; set; }

        private byte[] Buffer { get; set; }

        private int BufferPosition { get; set; }

        private int BufferLength { get; set; }

        private int MaxCharBytes { get; }

        private char Separator { get; }

        public List<string> ReadRow()
        {
            var result = new List<string>();
            while (true)
            {
                var value = ReadField(out var terminator, out var terminatorBytesUsed);
                if (value == null && terminator == '\0' && result.Count == 0)
                {
                    return null;
                }

                result.Add(value);

                if (terminator == '\r')
                {
                    // check for \r\n
                    Advance(terminatorBytesUsed);
                    var c = PeekChar(out terminatorBytesUsed);
                    if (c == '\n')
                    {
                        Advance(terminatorBytesUsed);
                    }

                    break;
                }
                else if (terminator == '\n')
                {
                    Advance(terminatorBytesUsed);
                    break;
                }
                else if (terminator == Separator)
                {
                    Advance(terminatorBytesUsed);
                    continue;
                }
                else if (terminator == '\0')
                {
                    break;
                }
                else
                {
                    throw new InvalidOperationException("Unexpected terminator " + terminator);
                }
            }

            return result;
        }

        public string ReadField(out char terminator, out int terminatorBytesUsed)
        {
            // Skip optional whitespace and quote character at the start
            char quoteChar = '\0';
            while (true)
            { 
                var c = PeekChar(out var bytesUsed);
                if (IsWhitespace(c))
                {
                    Advance(bytesUsed);
                    continue;
                }
                else if (c == '"' || c == '\'')
                {
                    quoteChar = c;
                    Advance(bytesUsed);
                    break;
                }
                else if (c == Separator || c == '\r' || c == '\n' || (c == '\0' && bytesUsed == 0))
                {
                    terminator = c;
                    terminatorBytesUsed = bytesUsed;
                    return null;
                }
                else
                {
                    break;
                }
            }

            // Positioned on first char, may be quote-terminated
            var result = new StringBuilder();
            var trailingWhitespaceCount = 0;
            while (true)
            {
                var c = PeekChar(out var bytesUsed);
                if (quoteChar == '\0')
                {
                    if (c == Separator || c == '\r' || c == '\n' || (c == '\0' && bytesUsed == 0))
                    {
                        terminator = c;
                        terminatorBytesUsed = bytesUsed;
                        break;
                    }
                    else
                    {
                        if (IsWhitespace(c))
                        {
                            trailingWhitespaceCount++;
                        }
                        else
                        {
                            trailingWhitespaceCount = 0;
                        }

                        result.Append(c);
                        Advance(bytesUsed);
                    }
                }
                else
                {
                    if (c == quoteChar)
                    {
                        Advance(bytesUsed);

                        c = PeekChar(out bytesUsed);
                        if (c == quoteChar)
                        {
                            // Escaped quote character
                            result.Append(quoteChar);
                            Advance(bytesUsed);
                        }
                        else
                        {
                            quoteChar = '\0';
                        }
                    }
                    else
                    {
                        result.Append(c);
                        Advance(bytesUsed);
                    }
                }
            }

            return result.ToString(0, result.Length - trailingWhitespaceCount);
        }

        private bool IsWhitespace(char c)
        {
            if (c == ' ')
            {
                return true;
            }

            if (Separator != '\t' && c == '\t')
            {
                return true;
            }

            return false;
        }

        private Encoding GetEncodingFromBom(byte[] bom, out int bomLength)
        {
            var encodings = new Encoding[]
            {
                Encoding.Unicode, Encoding.BigEndianUnicode, Encoding.UTF8
            };

            foreach (var encoding in encodings)
            {
                if (IsEncodingPreamble(bom, encoding, out int length))
                {
                    bomLength = length;
                    return encoding;
                }
            }

            bomLength = 0;
            return null;
        }

        private bool IsEncodingPreamble(byte[] bom, Encoding encoding, out int bomLength)
        {
            bomLength = 0;
            var preabmle = encoding.GetPreamble();
            if (preabmle.Length > bom.Length)
                return false;
            var i = 0;
            for (; i < preabmle.Length; i++)
            {
                if (preabmle[i] != bom[i])
                    return false;
            }

            bomLength = i;
            return true;
        }

        private char PeekChar(out int bytesUsed)
        {
            if (BufferPosition > BufferLength - MaxCharBytes && BaseStream.Position < BaseStream.Length)
            {
                throw new InvalidOperationException("Cannot peek past the buffer");
            }

            if (BufferPosition >= BufferLength)
            {
                bytesUsed = 0;
                return '\0';
            }

            var c = new char[1];
            Decoder.Convert(Buffer, BufferPosition, MaxCharBytes, c, 0, 1, true, out bytesUsed, out var charsUsed, out var completed);
            return c[0];
        }
        
        private void Advance(int count)
        {
            BufferPosition += count;

            if (BufferLength - BufferPosition < MaxCharBytes && BaseStream.Position < BaseStream.Length)
            {
                ReadBuffer();
            }
        }

        private void ReadBuffer()
        {
            var remainingBytes = BufferLength - BufferPosition;
            Array.Copy(Buffer, BufferPosition, Buffer, 0, remainingBytes);

            var bytesRead = BaseStream.Read(Buffer, remainingBytes, BufferSize - remainingBytes);
            BufferPosition = 0;
            BufferLength = bytesRead + remainingBytes;
        }
    }
}
