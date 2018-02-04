using System;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Helper class for parsing the BIFF8 Shared String Table (SST)
    /// </summary>
    internal class XlsSSTReader
    {
        public XlsSSTReader(XlsBiffSST sst, XlsBiffStream biffStream)
        {
            Sst = sst;
            BiffStream = biffStream;
            CurrentRecord = Sst;
            CurrentRecordOffset = 4 + 8; // +4 skips BIFF header, +8 skips SST header
        }

        private XlsBiffSST Sst { get; }

        private XlsBiffStream BiffStream { get; }

        private XlsBiffRecord CurrentRecord { get; set; }

        /// <summary>
        /// Gets or sets the offset into the current record's byte content. May point at the end when the current record has been parsed entirely.
        /// </summary>
        private int CurrentRecordOffset { get; set; }

        /// <summary>
        /// Reads an SST string potentially spanning multiple records
        /// </summary>
        /// <returns>The string</returns>
        public IXlsString ReadString()
        {
            EnsureRecord();

            var header = new XlsSSTStringHeader(CurrentRecord.Bytes, (uint)(CurrentRecord.Offset + CurrentRecordOffset));
            Advance((int)header.HeadSize);

            var remainingCharacters = (int)header.CharacterCount;

            const int XlsUnicodeStringHeaderSize = 3;

            byte[] result = new byte[XlsUnicodeStringHeaderSize + remainingCharacters * 2];
            result[0] = (byte)(remainingCharacters & 0x00FF);
            result[1] = (byte)((remainingCharacters & 0xFF00) >> 8);
            result[2] = 1; // IsMultiByte = true

            var resultOffset = XlsUnicodeStringHeaderSize;

            bool isMultiByte = header.IsMultiByte;
            while (remainingCharacters > 0)
            {
                if (EnsureRecord())
                {
                    // Continue records for string data start with a multibyte header
                    var b = ReadByte();
                    isMultiByte = b != 0;
                }

                var bytesPerCharacter = isMultiByte ? 2 : 1;
                var maxRecordCharacters = (CurrentRecord.Size - CurrentRecordOffset) / bytesPerCharacter;
                var readCharacters = Math.Min(maxRecordCharacters, remainingCharacters);
                
                ReadUnicodeBytes(result, resultOffset, readCharacters, isMultiByte);

                resultOffset += readCharacters * 2; // The result is always multibyte
                remainingCharacters -= readCharacters;
            }

            // Skip formatting runs and phonetic/extended data. Can also span
            // multiple Continue records
            Advance((int)header.TailSize);

            return new XlsUnicodeString(result, 0);
        }

        private void ReadUnicodeBytes(byte[] dest, int offset, int characterCount, bool isMultiByte)
        {
            if (CurrentRecordOffset >= CurrentRecord.Size)
            {
                throw new InvalidOperationException("SST read position out of range");
            }

            if (characterCount == 0)
            {
                throw new InvalidOperationException("Bad SST format");
            }

            if (isMultiByte)
            {
                Array.Copy(CurrentRecord.Bytes, CurrentRecord.Offset + CurrentRecordOffset, dest, offset, characterCount * 2);
                CurrentRecordOffset += characterCount * 2;
            }
            else
            {
                for (int i = 0; i < characterCount; i++)
                {
                    dest[offset + i * 2] = CurrentRecord.Bytes[CurrentRecord.Offset + CurrentRecordOffset + i];
                    dest[offset + i * 2 + 1] = 0;
                }

                CurrentRecordOffset += characterCount;
            }
        }

        private byte ReadByte()
        {
            if (CurrentRecordOffset >= CurrentRecord.Size)
            {
                throw new InvalidOperationException("SST read position out of range");
            }

            var result = CurrentRecord.Bytes[CurrentRecord.Offset + CurrentRecordOffset];
            CurrentRecordOffset++;
            return result;
        }

        /// <summary>
        /// If the read position is exactly at the end of a record:
        /// Read the next continue record and update the read position.
        /// </summary>
        private bool EnsureRecord()
        {
            if (CurrentRecordOffset == CurrentRecord.Size)
            {
                CurrentRecord = BiffStream.Read();
                if (CurrentRecord == null || CurrentRecord.Id != BIFFRECORDTYPE.CONTINUE)
                {
                    throw new InvalidOperationException("Bad SST format");
                }

                CurrentRecordOffset = 4; // +4 skips BIFF header
                return true;
            }

            return false;
        }

        /// <summary>
        /// Advances the read position a number of bytes, potentially spanning
        /// multiple records.
        /// NOTE: If the new read position ends on a record boundary, 
        /// the next record will not be read, and the read position will point
        /// at the end of the record! Must call EnsureRecord() as needed
        /// to read the next continue record and reset the read position. 
        /// </summary>
        /// <param name="bytes">Number of bytes to skip</param>
        private void Advance(int bytes)
        {
            var size = CurrentRecord.Size;

            while (CurrentRecordOffset + bytes > size)
            {
                bytes = Math.Min((CurrentRecordOffset + bytes) - size, bytes);

                CurrentRecordOffset = CurrentRecord.Size;
                EnsureRecord();
                size = CurrentRecord.Size;
            }

            CurrentRecordOffset += bytes;
        }
    }
}
