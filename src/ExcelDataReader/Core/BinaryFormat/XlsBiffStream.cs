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
    private byte[] _headerBuffer = new byte[4];

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
    public int Position { get => (int)BaseStream.Position; set => Seek(value, SeekOrigin.Begin); }

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
        var recordOffset = (int)stream.Position;
        stream.ReadAtLeast(_headerBuffer, 0, 4);

        // Does this work on a big endian system?
        var id = (BIFFRECORDTYPE)BitConverter.ToUInt16(_headerBuffer, 0);
        ushort recordSize = BitConverter.ToUInt16(_headerBuffer, 2);

#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        var bytes = System.Buffers.ArrayPool<byte>.Shared.Rent(4 + recordSize);
#else
        var bytes = new byte[4 + recordSize];
#endif
        Array.Copy(_headerBuffer, bytes, 4);
        stream.ReadAtLeast(bytes, 4, recordSize);
        
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
            case BIFFRECORDTYPE.BLANK:
            case BIFFRECORDTYPE.BLANK_OLD:
                return new XlsBiffBlankCell(bytes);
            case BIFFRECORDTYPE.MULBLANK:
                return new XlsBiffMulBlankCell(bytes);
            case BIFFRECORDTYPE.LABEL_OLD:
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
                return new XlsBiffWindow1(bytes);
            case BIFFRECORDTYPE.CODEPAGE:
                return new XlsBiffSimpleValueRecord(bytes);
            case BIFFRECORDTYPE.FNGROUPCOUNT:
                return new XlsBiffSimpleValueRecord(bytes);
            case BIFFRECORDTYPE.RECORD1904:
                return new XlsBiffSimpleValueRecord(bytes);
            case BIFFRECORDTYPE.BOOKBOOL:
                return new XlsBiffSimpleValueRecord(bytes);
            case BIFFRECORDTYPE.BACKUP:
                return new XlsBiffSimpleValueRecord(bytes);
            case BIFFRECORDTYPE.HIDEOBJ:
                return new XlsBiffSimpleValueRecord(bytes);
            case BIFFRECORDTYPE.USESELFS:
                return new XlsBiffSimpleValueRecord(bytes);
            case BIFFRECORDTYPE.UNCALCED:
                return new XlsBiffUncalced(bytes);
            case BIFFRECORDTYPE.QUICKTIP:
                return new XlsBiffQuickTip(bytes);
            case BIFFRECORDTYPE.MSODRAWING:
                return new XlsBiffMSODrawing(bytes);
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
        var bytes = new byte[blockOffset];
        CryptoHelpers.DecryptBytes(CipherTransform, bytes);
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

#if NETSTANDARD2_1_OR_GREATER
            var block = System.Buffers.ArrayPool<byte>.Shared.Rent(chunkSize);
#else
            var block = new byte[chunkSize];
#endif

            Array.Copy(bytes, position, block, 0, chunkSize);

            var decryptedblock = CryptoHelpers.DecryptBytes(CipherTransform, block);
            for (var i = 0; i < decryptedblock.Length; i++)
            {
                if (position >= startDecrypt)
                    bytes[position] = decryptedblock[i];
                position++;
            }
        }
    }
}
