namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Assembles BIFF8 strings across borrowed SST/STRING and CONTINUE frames.
/// Only copied character data and header/tail progress survive a call.
/// </summary>
internal sealed class XlsSSTReader
{
    private readonly byte[] _header = new byte[9];
    private int _headerBytes;
    private int _headerSize = 3;
    private int _characterCount;
    private int _remainingCharacters;
    private uint _remainingTail;
    private bool _isMultiByte;
    private int _resultOffset;
    private StringState _state;

    private enum StringState
    {
        Header,
        Data,
        Tail,
    }

    private byte[] CurrentResult { get; set; } = [];

    public void ReadStrings(XlsBiffFrame frame, int offset, bool continuation, ISharedStringSink sink, int maxStrings)
    {
        int end = frame.RecordSize + 4;
        if (offset == end || maxStrings == 0)
            return;

        if (continuation && _state == StringState.Data)
            _isMultiByte = frame.ReadByteAt(offset++) != 0;

        int stringsRead = 0;
        while (stringsRead < maxStrings)
        {
            if (_state == StringState.Header)
            {
                if (!ReadHeader(frame, ref offset, end))
                    return;

                _characterCount = BitConverter.ToUInt16(_header, 0);
                _remainingCharacters = _characterCount;
                _isMultiByte = (_header[2] & 1) != 0;
                bool formatted = (_header[2] & 8) != 0;
                _remainingTail = formatted ? (uint)BitConverter.ToUInt16(_header, 3) * 4 : 0;
                if ((_header[2] & 4) != 0)
                    _remainingTail += BitConverter.ToUInt32(_header, formatted ? 5 : 3);

                int size = _characterCount * 2;
                if (CurrentResult.Length < size)
                    CurrentResult = new byte[size];
                _resultOffset = 0;
                _state = StringState.Data;
                if (offset == end && _remainingCharacters > 0)
                    return;
            }

            if (_state == StringState.Data)
            {
                int characters = Math.Min((end - offset) / (_isMultiByte ? 2 : 1), _remainingCharacters);
                if (_isMultiByte)
                {
                    Buffer.BlockCopy(frame.Bytes, offset, CurrentResult, _resultOffset, characters * 2);
                    offset += characters * 2;
                }
                else
                {
                    for (int i = 0; i < characters; i++)
                    {
                        CurrentResult[_resultOffset + i * 2] = frame.Bytes[offset + i];
                        CurrentResult[_resultOffset + i * 2 + 1] = 0;
                    }

                    offset += characters;
                }

                _resultOffset += characters * 2;
                _remainingCharacters -= characters;
                if (_isMultiByte && _remainingCharacters > 0 && end - offset == 1)
                    offset++;
                if (_remainingCharacters > 0)
                    return;

                _state = StringState.Tail;
            }

            if (_state == StringState.Tail)
            {
                int skip = (int)Math.Min((uint)(end - offset), _remainingTail);
                offset += skip;
                _remainingTail -= (uint)skip;
                if (_remainingTail != 0)
                    return;

                AddCurrentString(sink);
                _state = StringState.Header;
                _headerBytes = 0;
                _headerSize = 3;
                stringsRead++;
            }
        }
    }

    public void Flush(ISharedStringSink sink)
    {
        // A missing formatting/phonetic tail does not invalidate the character data.
        if (_state == StringState.Tail)
            AddCurrentString(sink);
        CurrentResult = [];
    }

    private bool ReadHeader(XlsBiffFrame frame, ref int offset, int end)
    {
        while (_headerBytes < _headerSize)
        {
            int count = Math.Min(end - offset, _headerSize - _headerBytes);
            Buffer.BlockCopy(frame.Bytes, offset, _header, _headerBytes, count);
            offset += count;
            _headerBytes += count;
            if (_headerBytes < _headerSize)
                return false;
            if (_headerSize == 3)
                _headerSize += ((_header[2] & 8) != 0 ? 2 : 0) + ((_header[2] & 4) != 0 ? 4 : 0);
        }

        return true;
    }

    private void AddCurrentString(ISharedStringSink sink) => sink.AddUtf16(CurrentResult, 0, _characterCount);
}