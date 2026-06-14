using System.Globalization;
using System.Security.Cryptography;
using ExcelDataReader.Core.OfficeCrypto;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents a BIFF stream.
/// </summary>
internal sealed class XlsBiffStream : IDisposable
{
    // Read-ahead buffer: a single BaseStream.Read() fills 1 024 bytes that then serve
    // ~70 sequential small records (avg 14 bytes each) without further CompoundStream calls.
    // Any Seek() sets _readAheadStart == _readAheadEnd to invalidate it.
    // Uses only byte[], Buffer.BlockCopy and Stream.Read — available on all target frameworks.
    private const int ReadAheadSize = 1024;

    private readonly byte[] _headerBuffer = new byte[4];

    private readonly byte[] _readAheadBuffer = new byte[ReadAheadSize];
    private int _readAheadStart;
    private int _readAheadEnd;

    public XlsBiffStream(Stream baseStream, int offset = 0, int explicitVersion = 0, BIFFTYPE? defaultType = null, string password = null, byte[] secretKey = null, EncryptionInfo encryption = null)
    {
        BaseStream = baseStream;
        Position = offset;

        var record = Read();
        if (record is XlsBiffBOF bof)
        {
            BiffVersion = explicitVersion == 0 ? XlsBiffStream.GetBiffVersion(bof) : explicitVersion;
            BiffType = bof.Type;

            if (secretKey == null)
                record = Read();
        }
        else if (explicitVersion > 0 && defaultType != null) 
        {
            BiffVersion = explicitVersion;
            BiffType = defaultType.Value;
        }

        CipherBlock = -1;
        if (secretKey != null)
        {
            SecretKey = secretKey;
            Encryption = encryption;
            Cipher = Encryption.CreateCipher();
        }
        else
        {
            var filePass = record as XlsBiffFilePass;
            filePass ??= Read() as XlsBiffFilePass;

            if (filePass != null)
            {
                Encryption = filePass.EncryptionInfo;

                if (Encryption.VerifyPassword("VelvetSweatshop"))
                {
                    // Magic password used for write-protected workbooks
                    password = "VelvetSweatshop";
                }
                else if (password == null || !Encryption.VerifyPassword(password))
                {
                    throw new InvalidPasswordException(Errors.ErrorInvalidPassword);
                }

                SecretKey = Encryption.GenerateSecretKey(password);
                Cipher = Encryption.CreateCipher();
            }
        }

        Position = offset;
    }

    public int BiffVersion { get; }

    public BIFFTYPE BiffType { get; }

    /// <summary>
    /// Gets the size of BIFF stream in bytes.
    /// </summary>
    public int Size => (int)BaseStream.Length;

    /// <summary>
    /// Gets or sets the current position in BIFF stream.
    /// </summary>
    public int Position
    {
        get => (int)BaseStream.Position - (_readAheadEnd - _readAheadStart);
        set => Seek(value, SeekOrigin.Begin);
    }

    public Stream BaseStream { get; }

    public byte[] SecretKey { get; }

    public EncryptionInfo Encryption { get; }

    public SymmetricAlgorithm Cipher { get; }

    /// <summary>
    /// Gets or sets the ICryptoTransform instance used to decrypt the current block.
    /// </summary>
    public ICryptoTransform CipherTransform { get; set; }

    /// <summary>
    /// Gets or sets the current block number being decrypted with CipherTransform.
    /// </summary>
    public int CipherBlock { get; set; }

    /// <summary>
    /// Sets stream pointer to the specified offset.
    /// </summary>
    /// <param name="offset">Offset value.</param>
    /// <param name="origin">Offset origin.</param>
    public void Seek(int offset, SeekOrigin origin)
    {
        // Discard buffered bytes: they belong to the old position.
        _readAheadStart = 0;
        _readAheadEnd = 0;
        BaseStream.Seek(offset, origin);

        if (Position < 0)
            throw new ArgumentOutOfRangeException(string.Format(CultureInfo.InvariantCulture, "{0} On offset={1}", Errors.ErrorBiffIlegalBefore, offset));
        if (Position > Size)
            throw new ArgumentOutOfRangeException(string.Format(CultureInfo.InvariantCulture, "{0} On offset={1}", Errors.ErrorBiffIlegalAfter, offset));

        if (SecretKey != null)
        { 
            CreateBlockDecryptor(offset / 1024);
            AlignBlockDecryptor(offset % 1024);
        }
    }

    /// <summary>
    /// Reads record under cursor and advances cursor position to next record.
    /// </summary>
    /// <returns>The record -or- null.</returns>
    public XlsBiffRecord Read()
    {
        // Minimum record size is 4
        if ((uint)Position + 4 >= Size)
            return null;

        var record = GetRecord(BaseStream);

        if (Position > Size)
        {
            record = null;
        }

        return record;
    }

    /// <summary>
    /// Returns record at specified offset.
    /// </summary>
    /// <param name="stream">The stream.</param>
    /// <returns>The record -or- null.</returns>
    public XlsBiffRecord GetRecord(Stream stream)
    {
        // Capture the logical record start before consuming header bytes from the read-ahead
        // buffer; this is the value DecryptRecord() needs for its block-number calculation.
        var recordOffset = Position;
        ReadFromReadAhead(_headerBuffer, 0, 4);

        // Does this work on a big endian system?
        var id = (BIFFRECORDTYPE)BitConverter.ToUInt16(_headerBuffer, 0);
        ushort recordSize = BitConverter.ToUInt16(_headerBuffer, 2);

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        // Records at or below SmallRecordPoolThreshold bytes use a plain heap allocation.
        // ArrayPool.Rent+Return overhead exceeds the benefit for these tiny buffers.
        // Rented buffers always have Length > SmallRecordPoolThreshold (pool rounds up to
        // the next power-of-two bucket), so XlsBiffRecord.Return() can tell them apart.
        var bytes = (4 + recordSize) <= XlsBiffRecord.SmallRecordPoolThreshold
            ? new byte[4 + recordSize]
            : System.Buffers.ArrayPool<byte>.Shared.Rent(4 + recordSize);
#else
        var bytes = new byte[4 + recordSize];
#endif
        Array.Copy(_headerBuffer, bytes, 4);
        ReadFromReadAhead(bytes, 4, recordSize);
        
        if (SecretKey != null)
            DecryptRecord(recordOffset, id, bytes, 4 + recordSize);

        int biffVersion = BiffVersion;

        switch (id)
        {
            case BIFFRECORDTYPE.BOF_V2:
            case BIFFRECORDTYPE.BOF_V3:
            case BIFFRECORDTYPE.BOF_V4:
            case BIFFRECORDTYPE.BOF:
                return new XlsBiffBOF(bytes);
            case BIFFRECORDTYPE.EOF:
                return new XlsBiffEof(bytes);
            case BIFFRECORDTYPE.INTERFACEHDR:
                return new XlsBiffInterfaceHdr(bytes);

            case BIFFRECORDTYPE.SST:
                return new XlsBiffSST(bytes);

            case BIFFRECORDTYPE.DEFAULTROWHEIGHT_V2:
            case BIFFRECORDTYPE.DEFAULTROWHEIGHT:
                return new XlsBiffDefaultRowHeight(bytes, biffVersion);
            case BIFFRECORDTYPE.ROW_V2:
            case BIFFRECORDTYPE.ROW:
                return new XlsBiffRow(bytes);

            case BIFFRECORDTYPE.BOOLERR:
            case BIFFRECORDTYPE.BOOLERR_OLD:
                return new XlsBiffBoolErrCell(bytes);
            case BIFFRECORDTYPE.BLANK:
            case BIFFRECORDTYPE.BLANK_OLD:
                return new XlsBiffBlankCell(bytes);
            case BIFFRECORDTYPE.MULBLANK:
                return new XlsBiffBlankCell(bytes);
            case BIFFRECORDTYPE.LABEL_V2:
            case BIFFRECORDTYPE.LABEL:
            case BIFFRECORDTYPE.RSTRING:
                return new XlsBiffLabelCell(bytes, biffVersion);
            case BIFFRECORDTYPE.LABELSST:
                return new XlsBiffLabelSSTCell(bytes);
            case BIFFRECORDTYPE.INTEGER:
            case BIFFRECORDTYPE.INTEGER_OLD:
                return new XlsBiffIntegerCell(bytes);
            case BIFFRECORDTYPE.NUMBER:
            case BIFFRECORDTYPE.NUMBER_OLD:
                return new XlsBiffNumberCell(bytes);
            case BIFFRECORDTYPE.RK:
                return new XlsBiffRKCell(bytes);
            case BIFFRECORDTYPE.MULRK:
                return new XlsBiffMulRKCell(bytes);
            case BIFFRECORDTYPE.FORMULA:
            case BIFFRECORDTYPE.FORMULA_V3:
            case BIFFRECORDTYPE.FORMULA_V4:
                return new XlsBiffFormulaCell(bytes, biffVersion);
            case BIFFRECORDTYPE.FORMAT_V23:
            case BIFFRECORDTYPE.FORMAT:
                return new XlsBiffFormatString(bytes, biffVersion);
            case BIFFRECORDTYPE.STRING:
            case BIFFRECORDTYPE.STRING_OLD:
                return new XlsBiffFormulaString(bytes, biffVersion);
            case BIFFRECORDTYPE.CONTINUE:
                return new XlsBiffContinue(bytes);
            case BIFFRECORDTYPE.DIMENSIONS:
            case BIFFRECORDTYPE.DIMENSIONS_V2 when bytes.Length >= 12:
                return new XlsBiffDimensions(bytes, biffVersion);
            case BIFFRECORDTYPE.BOUNDSHEET:
                return new XlsBiffBoundSheet(bytes, biffVersion);
            case BIFFRECORDTYPE.WINDOW1:
                return new XlsBiffRecord(bytes);
            case BIFFRECORDTYPE.CODEPAGE:
            case BIFFRECORDTYPE.FNGROUPCOUNT:
            case BIFFRECORDTYPE.DATE1904:
            case BIFFRECORDTYPE.BOOKBOOL:
            case BIFFRECORDTYPE.BACKUP:
            case BIFFRECORDTYPE.HIDEOBJ:
            case BIFFRECORDTYPE.USESELFS:
                return new XlsBiffSimpleValueRecord(bytes);
            case BIFFRECORDTYPE.FILEPASS:
                return new XlsBiffFilePass(bytes, biffVersion);
            case BIFFRECORDTYPE.HEADER:
            case BIFFRECORDTYPE.FOOTER:
                return new XlsBiffHeaderFooterString(bytes, biffVersion);
            case BIFFRECORDTYPE.CODENAME:
                return new XlsBiffCodeName(bytes);
            case BIFFRECORDTYPE.XF:
            case BIFFRECORDTYPE.XF_V2:
            case BIFFRECORDTYPE.XF_V3:
            case BIFFRECORDTYPE.XF_V4:
                return new XlsBiffXF(bytes, biffVersion);
            case BIFFRECORDTYPE.FONT:
                return new XlsBiffFont(bytes, biffVersion);
            case BIFFRECORDTYPE.MERGECELLS:
                return new XlsBiffMergeCells(bytes);
            case BIFFRECORDTYPE.COLINFO:
                return new XlsBiffColInfo(bytes);
            default:
                return new XlsBiffRecord(bytes);
        }
    }

    public void Dispose()
    {
        CipherTransform?.Dispose();
        ((IDisposable)Cipher)?.Dispose();
    }

    private static int GetBiffVersion(XlsBiffBOF bof)
    {
        switch (bof.Id)
        {
            case BIFFRECORDTYPE.BOF_V2:
                return 2;
            case BIFFRECORDTYPE.BOF_V3:
                return 3;
            case BIFFRECORDTYPE.BOF_V4:
                return 4;
            case BIFFRECORDTYPE.BOF:
                if (bof.Version == 0x200)
                    return 2;
                else if (bof.Version == 0x300)
                    return 3;
                else if (bof.Version == 0x400)
                    return 4;
                else if (bof.Version == 0x500 || bof.Version == 0)
                    return 5;
                if (bof.Version == 0x600)
                    return 8;
                break;
        }

        return 0;
    }

    // Serves 'count' bytes into dest[destOffset..], draining the read-ahead buffer first.
    // When the buffer is exhausted it is refilled with a single BaseStream.Read(1024).
    // For large bodies that exceed the remaining buffer the tail is read directly from
    // BaseStream in one call (avoids per-1024-chunk overhead for SST / other big records).
    private void ReadFromReadAhead(byte[] dest, int destOffset, int count)
    {
        // Fast path: serve from the in-memory buffer when possible.
        int available = _readAheadEnd - _readAheadStart;
        if (available > 0)
        {
            int fromBuffer = Math.Min(available, count);
            Buffer.BlockCopy(_readAheadBuffer, _readAheadStart, dest, destOffset, fromBuffer);
            _readAheadStart += fromBuffer;

            if (fromBuffer == count)
                return;

            destOffset += fromBuffer;
            count -= fromBuffer;
        }

        // Buffer was exhausted before satisfying the request.
        // For small leftovers refill the buffer and serve from it;
        // for large reads go directly to BaseStream to avoid per-chunk overhead.
        if (count <= ReadAheadSize)
        {
            _readAheadStart = 0;
            _readAheadEnd = BaseStream.Read(_readAheadBuffer, 0, ReadAheadSize);
            if (_readAheadEnd < count)
                throw new EndOfStreamException();

            Buffer.BlockCopy(_readAheadBuffer, 0, dest, destOffset, count);
            _readAheadStart = count;
        }
        else
        {
            // Large record body: read directly into the caller's buffer.
            // The read-ahead buffer remains empty (_readAheadStart == _readAheadEnd == 0)
            // and will be refilled on the next call.
            _readAheadStart = 0;
            _readAheadEnd = 0;
            BaseStream.ReadAtLeast(dest, destOffset, count);
        }
    }

    /// <summary>
    /// Create an ICryptoTransform instance to decrypt a 1024-byte block.
    /// </summary>
    private void CreateBlockDecryptor(int blockNumber)
    {
        CipherTransform?.Dispose();

        var blockKey = Encryption.GenerateBlockKey(blockNumber, SecretKey);
        CipherTransform = Cipher.CreateDecryptor(blockKey, null);
        CipherBlock = blockNumber;
    }

    /// <summary>
    /// Decrypt some dummy bytes to align the decryptor with the position in the current 1024-byte block.
    /// </summary>
    private void AlignBlockDecryptor(int blockOffset)
    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        var bytes = System.Buffers.ArrayPool<byte>.Shared.Rent(blockOffset);
        try
        {
            CryptoHelpers.DecryptBytes(CipherTransform, bytes, blockOffset);
        }
        finally
        {
            System.Buffers.ArrayPool<byte>.Shared.Return(bytes);
        }
#else
        var bytes = new byte[blockOffset];
        CryptoHelpers.DecryptBytes(CipherTransform, bytes, blockOffset);
#endif
    }

    private void DecryptRecord(int startPosition, BIFFRECORDTYPE id, byte[] bytes, int recordSize)
    {
        // Decrypt the last read record, find it's start offset relative to the current stream position
        int startDecrypt = 4;
        switch (id)
        {
            case BIFFRECORDTYPE.BOF:
            case BIFFRECORDTYPE.FILEPASS:
            case BIFFRECORDTYPE.INTERFACEHDR:
                startDecrypt = recordSize;
                break;
            case BIFFRECORDTYPE.BOUNDSHEET:
                startDecrypt += 4; // For some reason the sheet offset is not encrypted
                break;
        }

        // Max chunk size per iteration is 1024 (one encryption block boundary).
        // Rent both buffers once and reuse across iterations to avoid per-iteration allocations.
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        var inputBlock = System.Buffers.ArrayPool<byte>.Shared.Rent(1024);
        var outputBlock = System.Buffers.ArrayPool<byte>.Shared.Rent(1024);
#else
        var inputBlock = new byte[1024];
        var outputBlock = new byte[1024];
#endif
        try
        {
            var position = 0;
            while (position < recordSize)
            {
                var offset = startPosition + position;
                int blockNumber = offset / 1024;
                var blockOffset = offset % 1024;

                if (blockNumber != CipherBlock)
                {
                    CreateBlockDecryptor(blockNumber);
                }

                if (Encryption.IsXor)
                {
                    // Bypass everything and hook into the XorTransform instance to set the XorArrayIndex pr record.
                    // This is a hack to use the XorTransform otherwise transparently to the other encryption methods.
                    var xorTransform = (XorManaged.XorTransform)CipherTransform;
                    xorTransform.XorArrayIndex = offset + recordSize - 4;
                }

                // Decrypt at most up to the next 1024 byte boundary
                var chunkSize = Math.Min(recordSize - position, 1024 - blockOffset);

                Array.Copy(bytes, position, inputBlock, 0, chunkSize);
                CryptoHelpers.DecryptBytes(CipherTransform, inputBlock, chunkSize, outputBlock);

                for (var i = 0; i < chunkSize; i++)
                {
                    if (position >= startDecrypt)
                        bytes[position] = outputBlock[i];
                    position++;
                }
            }
        }
        finally
        {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
            System.Buffers.ArrayPool<byte>.Shared.Return(inputBlock);
            System.Buffers.ArrayPool<byte>.Shared.Return(outputBlock);
#endif
        }
    }
}
