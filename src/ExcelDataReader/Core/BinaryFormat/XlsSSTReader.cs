using ExcelDataReader.Core;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Helper class for parsing the BIFF8 Shared String Table (SST).
/// </summary>
internal sealed class XlsSSTReader
{
    private enum SstState
    {
        StartStringHeader,
        StringHeader,
        StringData,
        StringTail,
    }

    private XlsBiffRecord CurrentRecord { get; set; } = null!;

    /// <summary>
    /// Gets or sets the offset into the current record's byte content. May point at the end when the current record has been parsed entirely.
    /// </summary>
    private int CurrentRecordOffset { get; set; }

    private SstState CurrentState { get; set; } = SstState.StartStringHeader;

    private XlsSSTStringHeader CurrentHeader { get; set; } = null!;

    private int CurrentRemainingCharacters { get; set; }

    private byte[] CurrentResult { get; set; } = null!;

    private int CurrentResultOffset { get; set; }

    private int CurrentHeaderBytes { get; set; }

    private int CurrentTailBytes { get; set; }

    private bool CurrentIsMultiByte { get; set; }

    public void ReadStringsFromSST(XlsBiffSST sst, ISharedStringSink sink, int maxStrings)
    {
        CurrentRecord = sst;
        CurrentRecordOffset = 4 + 8;
        ReadStringsToSink(sink, maxStrings);
    }

    public void ReadStringsFromContinue(XlsBiffContinue sstContinue, ISharedStringSink sink, int maxStrings)
    {
        CurrentRecord = sstContinue;
        CurrentRecordOffset = 4;

        if (sstContinue.Size - CurrentRecordOffset == 0)
            return;

        if (CurrentState == SstState.StringData)
            CurrentIsMultiByte = ReadByte() != 0;

        ReadStringsToSink(sink, maxStrings);
    }

    public void Flush(ISharedStringSink sink)
    {
        if (CurrentState == SstState.StringTail)
            AddCurrentStringToSink(sink);

        CurrentResult = [];
    }

    private bool TryReadString(ISharedStringSink sink) => TryReadStringCore(sink);

    private void ReadStringsToSink(ISharedStringSink sink, int maxStrings)
    {
        int stringsRead = 0;
        while (stringsRead < maxStrings && TryReadString(sink))
        {
            stringsRead++;
        }
    }

    private bool TryReadStringCore(ISharedStringSink sink)
    {
        if (CurrentState == SstState.StartStringHeader)
        {
            if (CurrentRecord.Size - CurrentRecordOffset == 0)
            {
                return false;
            }

            CurrentHeader = new XlsSSTStringHeader(CurrentRecord.Bytes, CurrentRecordOffset);
            CurrentIsMultiByte = CurrentHeader.IsMultiByte;
            CurrentHeaderBytes = (int)CurrentHeader.HeadSize;
            CurrentRemainingCharacters = CurrentHeader.CharacterCount;

            const int XlsUnicodeStringHeaderSize = 3;

            int size = XlsUnicodeStringHeaderSize + CurrentRemainingCharacters * 2;
            if (CurrentResult == null || CurrentResult.Length < size)
                CurrentResult = new byte[size];
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
                return false;
            }

            CurrentState = SstState.StringData;

            if (CurrentRecord.Size - CurrentRecordOffset == 0)
            {
                // End of buffer before string data. Return false in StringData state to ensure reading the multibyte flag in the next record
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
                CurrentTailBytes -= advanceBytes;
                return false;
            }

            CurrentState = SstState.StartStringHeader;
            AddCurrentStringToSink(sink);
            return true;
        }

        throw new InvalidOperationException("Unexpected state in SST reader");
    }

    private void AddCurrentStringToSink(ISharedStringSink sink)
    {
        int end = 3 + CurrentHeader.CharacterCount * 2;
        if (CurrentResultOffset < end)
            Array.Clear(CurrentResult, CurrentResultOffset, end - CurrentResultOffset);
        sink.AddUtf16(CurrentResult, 3, CurrentHeader.CharacterCount);
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
