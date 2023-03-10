using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using ExcelDataReader.Core.OpenXmlFormat.Records;

#nullable enable

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat
{
    internal abstract class BiffReader : RecordReader
    {
        private readonly byte[] _buffer = new byte[128];

        public BiffReader(Stream stream)
        {
            Stream = stream ?? throw new ArgumentNullException(nameof(stream));
        }

        protected Stream Stream { get; }

        public override Record? Read()
        {
            if (!TryReadVariableValue(out var recordId) ||
                !TryReadVariableValue(out var recordLength))
                return null;

            byte[] buffer = recordLength < _buffer.Length ? _buffer : new byte[recordLength];
            if (Stream.Read(buffer, 0, (int)recordLength) != recordLength)
                return null;

            return ReadOverride(buffer, recordId, recordLength);
        }

        protected static uint GetDWord(byte[] buffer, uint offset)
        {
            uint result = (uint)buffer[offset + 3] << 24;
            result += (uint)buffer[offset + 2] << 16;
            result += (uint)buffer[offset + 1] << 8;
            result += buffer[offset];
            return result;
        }

        protected static int GetInt32(byte[] buffer, uint offset)
        {
            int result = buffer[offset + 3] << 24;
            result += buffer[offset + 2] << 16;
            result += buffer[offset + 1] << 8;
            result += buffer[offset];
            return result;
        }

        protected static ushort GetWord(byte[] buffer, uint offset)
        {
            ushort result = (ushort)(buffer[offset + 1] << 8);
            result += buffer[offset];
            return result;
        }

        protected static string GetString(byte[] buffer, uint offset, uint length)
        {
            StringBuilder sb = new StringBuilder((int)length);
            for (uint i = offset; i < offset + 2 * length; i += 2)
                sb.Append((char)GetWord(buffer, i));
            return sb.ToString();
        }

        protected static string? GetNullableString(byte[] buffer, ref uint offset)
        {
            var length = GetDWord(buffer, offset);
            offset += 4;
            if (length == uint.MaxValue)
                return null;
            StringBuilder sb = new StringBuilder((int)length);
            uint end = offset + length * 2;
            for (; offset < end; offset += 2)
                sb.Append((char)GetWord(buffer, offset));
            return sb.ToString();
        }

        protected static double GetRkNumber(byte[] buffer, uint offset)
        {
            double result;

            byte flags = buffer[offset];

            if ((flags & 0x02) != 0)
            {
                result = GetInt32(buffer, offset) >> 2;
            }
            else
            {
                result = BitConverter.Int64BitsToDouble((GetDWord(buffer, offset) & -4) << 32);
            }

            if ((flags & 0x01) != 0)
            {
                result /= 100;
            }

            return result;
        }

        protected static double GetDouble(byte[] buffer, uint offset)
        {
            uint num = GetDWord(buffer, offset);
            uint num2 = GetDWord(buffer, offset + 4);
            long num3 = ((long)num2 << 32) | num;
            return BitConverter.Int64BitsToDouble(num3);
        }

        protected abstract Record ReadOverride(byte[] buffer, uint recordId, uint recordLength);

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                Stream.Dispose();
        }

        private bool TryReadVariableValue(out uint value)
        {
            value = 0;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;

            byte b1 = _buffer[0];
            value = (uint)(b1 & 0x7F);

            if ((b1 & 0x80) == 0)
                return true;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;
            byte b2 = _buffer[0];
            value = ((uint)(b2 & 0x7F) << 7) | value;

            if ((b2 & 0x80) == 0)
                return true;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;
            byte b3 = _buffer[0];
            value = ((uint)(b3 & 0x7F) << 14) | value;

            if ((b3 & 0x80) == 0)
                return true;

            if (Stream.Read(_buffer, 0, 1) == 0)
                return false;
            byte b4 = _buffer[0];
            value = ((uint)(b4 & 0x7F) << 21) | value;

            return true;
        }        
    }
}
