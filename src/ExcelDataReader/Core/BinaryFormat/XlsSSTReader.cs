using System;
using System.Collections.Generic;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Helper class for parsing the BIFF8 Shared String Table (SST)
    /// </summary>
    internal class XlsSSTReader
    {
        private enum SstState
        {
            StartStringHeader,
            StringHeader,
            StringData,
            StringTail,
        }

        private XlsBiffRecord CurrentRecord { get; set; }

        /// <summary>
        /// Gets or sets the offset into the current record's byte content. May point at the end when the current record has been parsed entirely.
        /// </summary>
        private int CurrentRecordOffset { get; set; }

        private SstState CurrentState { get; set; } = SstState.StartStringHeader;

        private XlsSSTStringHeader CurrentHeader { get; set; }

        private int CurrentRemainingCharacters { get; set; }

        private byte[] CurrentResult { get; set; }

        private int CurrentResultOffset { get; set; }

        private int CurrentHeaderBytes { get; set; }

        private int CurrentTailBytes { get; set; }

        private bool CurrentIsMultiByte { get; set; } = false;

        public IEnumerable<IXlsString> ReadStringsFromSST(XlsBiffSST sst)
        {
            CurrentRecord = sst;
            CurrentRecordOffset = 4 + 8;

            while (true)
            {
                if (!TryReadString(out var result))
                {
                    yield break;
                }

                yield return result;
            }
        }

        public IEnumerable<IXlsString> ReadStringsFromContinue(XlsBiffContinue sstContinue)
        {
            CurrentRecord = sstContinue;
            CurrentRecordOffset = 4; // +4 skips BIFF header

            if (sstContinue.Size - CurrentRecordOffset == 0)
            {
                yield break;
            }

            if (CurrentState == SstState.StringData)
            {
                byte b = ReadByte();
                CurrentIsMultiByte = b != 0;
            }

            while (true)
            {
                if (!TryReadString(out var result))
                {
                    yield break;
                }

                yield return result;
            }
        }

        public IXlsString Flush()
        {
            if (CurrentState == SstState.StringTail)
            {
                return new XlsUnicodeString(CurrentResult, 0);
            }

            return null;
        }

        private bool TryReadString(out IXlsString result)
        {
            if (CurrentState == SstState.StartStringHeader)
            {
                if (CurrentRecord.Size - CurrentRecordOffset == 0)
                {
                    result = null;
                    return false;
                }

                CurrentHeader = new XlsSSTStringHeader(CurrentRecord.Bytes, CurrentRecordOffset);
                CurrentIsMultiByte = CurrentHeader.IsMultiByte;
                CurrentHeaderBytes = (int)CurrentHeader.HeadSize;
                CurrentRemainingCharacters = (int)CurrentHeader.CharacterCount;

                const int XlsUnicodeStringHeaderSize = 3;

                CurrentResult = new byte[XlsUnicodeStringHeaderSize + CurrentRemainingCharacters * 2];
                CurrentResult[0] = (byte)(CurrentRemainingCharacters & 0x00FF);
                CurrentResult[1] = (byte)((CurrentRemainingCharacters & 0xFF00) >> 8);
                CurrentResult[2] = 1; // IsMultiByte = true

                CurrentResultOffset = XlsUnicodeStringHeaderSize;

                CurrentState = SstState.StringHeader;
            }

            if (CurrentState == SstState.StringHeader)
            {
                if (!Advance(CurrentHeaderBytes, out int advanceBytes))
                {
                    CurrentHeaderBytes -= advanceBytes;
                    result = null;
                    return false;
                }

                CurrentState = SstState.StringData;

                if (CurrentRecord.Size - CurrentRecordOffset == 0)
                {
                    // End of buffer before string data. Return false in StringData state to ensure reading the multibyte flag in the next record
                    result = null;
                    return false;
                }
            }

            if (CurrentState == SstState.StringData)
            {
                var bytesPerCharacter = CurrentIsMultiByte ? 2 : 1;
                var maxRecordCharacters = (CurrentRecord.Size - CurrentRecordOffset) / bytesPerCharacter;
                var readCharacters = Math.Min(maxRecordCharacters, CurrentRemainingCharacters);

                ReadUnicodeBytes(CurrentResult, CurrentResultOffset, readCharacters, CurrentIsMultiByte);

                CurrentResultOffset += readCharacters * 2; // The result is always multibyte
                CurrentRemainingCharacters -= readCharacters;

                if (CurrentIsMultiByte && CurrentRecord.Size - CurrentRecordOffset == 1)
                {
                    // Skip leftover byte at the end of a multibyte Continue record
                    ReadByte();
                }

                if (CurrentRemainingCharacters > 0 && CurrentRecord.Size - CurrentRecordOffset == 0)
                {
                    result = null;
                    return false;
                }

                CurrentState = SstState.StringTail;
                CurrentTailBytes = (int)CurrentHeader.TailSize;
            }

            if (CurrentState == SstState.StringTail)
            {
                // Skip formatting runs and phonetic/extended data. Can also span
                // multiple Continue records
                if (!Advance(CurrentTailBytes, out var advanceBytes))
                {
                    result = null;
                    CurrentTailBytes -= advanceBytes;
                    return false;
                }

                CurrentState = SstState.StartStringHeader;
                result = new XlsUnicodeString(CurrentResult, 0);
                return true;
            }

            throw new InvalidOperationException("Unexpected state in SST reader");
        }

        private void ReadUnicodeBytes(byte[] dest, int offset, int characterCount, bool isMultiByte)
        {
            if (isMultiByte)
            {
                Array.Copy(CurrentRecord.Bytes, CurrentRecordOffset, dest, offset, characterCount * 2);
                CurrentRecordOffset += characterCount * 2;
            }
            else
            {
                for (int i = 0; i < characterCount; i++)
                {
                    dest[offset + i * 2] = CurrentRecord.Bytes[CurrentRecordOffset + i];
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

            var result = CurrentRecord.Bytes[CurrentRecordOffset];
            CurrentRecordOffset++;
            return result;
        }

        private bool Advance(int bytes, out int advanceBytes)
        {
            advanceBytes = Math.Min(CurrentRecord.Size - CurrentRecordOffset, bytes);
            CurrentRecordOffset += advanceBytes;
            return bytes == advanceBytes;
        }
    }
}
