using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.CsvFormat
{
    /// <summary>
    /// Low level, reentrant CSV parser. Call ParseBuffer() in a loop, and finally Flush() to empty the internal buffers.
    /// </summary>
    internal class CsvParser
    {
        public CsvParser(char separator, Encoding encoding)
        {
            Separator = separator;

            Decoder = encoding.GetDecoder();
            Decoder.Fallback = new DecoderExceptionFallback();

            MaxCharBytes = encoding.GetMaxByteCount(1);
            Buffer = new byte[MaxCharBytes];

            State = CsvState.PreValue;
        }

        private enum CsvState
        {
            PreValue,
            Value,
            QuotedValue,
            QuotedValueQuote,
            Separator,
            Linebreak,
            EndOfFile,
        }

        private CsvState State { get; set; }

        private char QuoteChar { get; set; }

        private int TrailingWhitespaceCount { get; set; }

        private Decoder Decoder { get; }

        private int MaxCharBytes { get; }

        private bool HasCarriageReturn { get; set; }

        private char Separator { get; }

        private byte[] Buffer { get; set; }

        private int BufferWritePosition { get; set; }

        private StringBuilder ValueResult { get; set; } = new StringBuilder();

        private List<string> RowResult { get; set; } = new List<string>();

        private List<List<string>> RowsResult { get; set; } = new List<List<string>>();

        public void ParseBuffer(byte[] bytes, int offset, int count, out List<List<string>> rows)
        {
            for (var i = 0; i < count; i++)
                ParseByte(bytes[offset + i]);

            rows = RowsResult;
            RowsResult = new List<List<string>>();
        }

        public void Flush(out List<List<string>> rows)
        {
            while (BufferWritePosition > 0)
            {
                DecodeChar();
            }

            if (State != CsvState.PreValue)
            {
                AddValueToRow();
                AddRowToResult();
            }

            rows = RowsResult;
            RowsResult = new List<List<string>>();
        }

        private void ParseByte(byte b)
        {
            Buffer[BufferWritePosition] = b;
            BufferWritePosition++;

            if (BufferWritePosition == MaxCharBytes)
            {
                DecodeChar();
            }
        }

        private void DecodeChar()
        {
            var c = new char[1];
            Decoder.Convert(Buffer, 0, BufferWritePosition, c, 0, 1, true, out var bytesUsed, out var charsUsed, out var completed);
            ParseChar(c[0], bytesUsed);

            Array.Copy(Buffer, bytesUsed, Buffer, 0, BufferWritePosition - bytesUsed);
            BufferWritePosition -= bytesUsed;
        }

        private void ParseChar(char c, int bytesUsed)
        {
            var parsed = false;
            while (!parsed)
            {
                switch (State)
                {
                    case CsvState.PreValue:
                        parsed = ReadPreValue(c, bytesUsed);
                        break;
                    case CsvState.Value:
                        parsed = ReadValue(c, bytesUsed);
                        break;
                    case CsvState.QuotedValue:
                        parsed = ReadQuotedValue(c, bytesUsed);
                        break;
                    case CsvState.QuotedValueQuote:
                        parsed = ReadQuotedValueQuote(c, bytesUsed);
                        break;
                    case CsvState.Separator:
                        parsed = ReadSeparator(c, bytesUsed);
                        break;
                    case CsvState.Linebreak:
                        parsed = ReadLinebreak(c, bytesUsed);
                        break;
                    default:
                        throw new InvalidOperationException("Unhandled parser state: " + State);
                }
            }
        }

        private bool ReadPreValue(char c, int bytesUsed)
        {
            if (IsWhitespace(c))
            {
                return true;
            }
            else if (c == '"' || c == '\'')
            {
                QuoteChar = c;
                State = CsvState.QuotedValue;
                return true;
            }
            else if (c == Separator)
            {
                State = CsvState.Separator;
                return false;
            }
            else if (c == '\r' || c == '\n' || (c == '\0' && bytesUsed == 0))
            {
                State = CsvState.Linebreak;
                return false;
            }
            else if (c == '\0' && bytesUsed == 0)
            {
                State = CsvState.EndOfFile;
                return false;
            }
            else
            {
                State = CsvState.Value;
                return false;
            }
        }

        private bool ReadValue(char c, int bytesUsed)
        {
            if (c == Separator)
            {
                State = CsvState.Separator;
                return false;
            }
            else if (c == '\r' || c == '\n')
            {
                State = CsvState.Linebreak;
                return false;
            }
            else if (c == '\0' && bytesUsed == 0)
            {
                State = CsvState.EndOfFile;
                return false;
            }
            else
            {
                if (IsWhitespace(c))
                {
                    TrailingWhitespaceCount++;
                }
                else
                {
                    TrailingWhitespaceCount = 0;
                }

                ValueResult.Append(c);
                return true;
            }
        }

        private bool ReadQuotedValue(char c, int bytesUsed)
        {
            if (c == QuoteChar)
            {
                State = CsvState.QuotedValueQuote;
                return true;
            }
            else
            {
                ValueResult.Append(c);
                return true;
            }
        }

        private bool ReadQuotedValueQuote(char c, int bytesUsed)
        {
            if (c == QuoteChar)
            {
                // Is escaped quote
                ValueResult.Append(c);
                State = CsvState.QuotedValue;
                return true;
            }
            else
            {
                // End of quote, read remainder of field as a regular value until separator
                QuoteChar = '\0';
                State = CsvState.Value;
                return false;
            }
        }

        private bool ReadSeparator(char c, int bytesUsed)
        {
            AddValueToRow();
            State = CsvState.PreValue;
            return true;
        }

        private bool ReadLinebreak(char c, int bytesUsed)
        {
            if (HasCarriageReturn)
            {
                HasCarriageReturn = false;
                AddValueToRow();
                AddRowToResult();
                State = CsvState.PreValue;
                return c == '\n';
            }
            else if (c == '\r')
            {
                HasCarriageReturn = true;
                return true;
            }
            else
            {
                AddValueToRow();
                AddRowToResult();
                State = CsvState.PreValue;
                return true;
            }
        }

        private void AddValueToRow()
        {
            RowResult.Add(ValueResult.ToString(0, ValueResult.Length - TrailingWhitespaceCount)); 
            ValueResult = new StringBuilder();
            TrailingWhitespaceCount = 0;
        }

        private void AddRowToResult()
        {
            RowsResult.Add(RowResult);
            RowResult = new List<string>();
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
    }
}
