using System;
using System.IO;
using System.Security.Cryptography;
using ExcelDataReader.Core.OfficeCrypto;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a BIFF stream
    /// </summary>
    internal class XlsBiffStream : IDisposable
    {
        public XlsBiffStream(Stream baseStream, int offset = 0, int explicitVersion = 0, string password = null, byte[] secretKey = null, EncryptionInfo encryption = null)
        {
            BaseStream = baseStream;
            Position = offset;

            var bof = Read() as XlsBiffBOF;
            if (bof != null)
            { 
                BiffVersion = explicitVersion == 0 ? GetBiffVersion(bof) : explicitVersion;
                BiffType = bof.Type;
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
                var filePass = Read() as XlsBiffFilePass;
                if (filePass == null)
                    filePass = Read() as XlsBiffFilePass;

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
        /// Gets the size of BIFF stream in bytes
        /// </summary>
        public int Size => (int)BaseStream.Length;

        /// <summary>
        /// Gets or sets the current position in BIFF stream
        /// </summary>
        public int Position { get => (int)BaseStream.Position; set => Seek(value, SeekOrigin.Begin); }

        public Stream BaseStream { get; }

        public byte[] SecretKey { get; }

        public EncryptionInfo Encryption { get; }

        public SymmetricAlgorithm Cipher { get; }

        /// <summary>
        /// Gets or sets the ICryptoTransform instance used to decrypt the current block
        /// </summary>
        public ICryptoTransform CipherTransform { get; set; }

        /// <summary>
        /// Gets or sets the current block number being decrypted with CipherTransform
        /// </summary>
        public int CipherBlock { get; set; }

        /// <summary>
        /// Sets stream pointer to the specified offset
        /// </summary>
        /// <param name="offset">Offset value</param>
        /// <param name="origin">Offset origin</param>
        public void Seek(int offset, SeekOrigin origin)
        {
            BaseStream.Seek(offset, origin);

            if (Position < 0)
                throw new ArgumentOutOfRangeException(string.Format("{0} On offset={1}", Errors.ErrorBiffIlegalBefore, offset));
            if (Position > Size)
                throw new ArgumentOutOfRangeException(string.Format("{0} On offset={1}", Errors.ErrorBiffIlegalAfter, offset));

            if (SecretKey != null)
            { 
                CreateBlockDecryptor(offset / 1024);
                AlignBlockDecryptor(offset % 1024);
            }
        }

        /// <summary>
        /// Reads record under cursor and advances cursor position to next record
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
        /// Returns record at specified offset
        /// </summary>
        /// <param name="stream">The stream</param>
        /// <returns>The record -or- null.</returns>
        public XlsBiffRecord GetRecord(Stream stream)
        {
            var recordOffset = (int)stream.Position;
            var header = new byte[4];
            stream.Read(header, 0, 4);

            var id = (BIFFRECORDTYPE)BitConverter.ToUInt16(header, 0);
            int recordSize = BitConverter.ToUInt16(header, 2);

            var bytes = new byte[4 + recordSize];
            Array.Copy(header, bytes, 4);
            stream.Read(bytes, 4, recordSize);

            if (SecretKey != null)
                DecryptRecord(recordOffset, id, bytes);

            uint offset = 0;
            int biffVersion = BiffVersion;

            switch ((BIFFRECORDTYPE)id)
            {
                case BIFFRECORDTYPE.BOF_V2:
                case BIFFRECORDTYPE.BOF_V3:
                case BIFFRECORDTYPE.BOF_V4:
                case BIFFRECORDTYPE.BOF:
                    return new XlsBiffBOF(bytes, offset);
                case BIFFRECORDTYPE.EOF:
                    return new XlsBiffEof(bytes, offset);
                case BIFFRECORDTYPE.INTERFACEHDR:
                    return new XlsBiffInterfaceHdr(bytes, offset);

                case BIFFRECORDTYPE.SST:
                    return new XlsBiffSST(bytes, offset);

                case BIFFRECORDTYPE.INDEX:
                    return new XlsBiffIndex(bytes, offset, biffVersion == 8);
                case BIFFRECORDTYPE.DEFAULTROWHEIGHT_V2:
                case BIFFRECORDTYPE.DEFAULTROWHEIGHT:
                    return new XlsBiffDefaultRowHeight(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.ROW_V2:
                case BIFFRECORDTYPE.ROW:
                    return new XlsBiffRow(bytes, offset);
                case BIFFRECORDTYPE.DBCELL:
                    return new XlsBiffDbCell(bytes, offset);

                case BIFFRECORDTYPE.BOOLERR:
                case BIFFRECORDTYPE.BOOLERR_OLD:
                case BIFFRECORDTYPE.BLANK:
                case BIFFRECORDTYPE.BLANK_OLD:
                    return new XlsBiffBlankCell(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.MULBLANK:
                    return new XlsBiffMulBlankCell(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.LABEL_OLD:
                case BIFFRECORDTYPE.LABEL:
                case BIFFRECORDTYPE.RSTRING:
                    return new XlsBiffLabelCell(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.LABELSST:
                    return new XlsBiffLabelSSTCell(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.INTEGER:
                case BIFFRECORDTYPE.INTEGER_OLD:
                    return new XlsBiffIntegerCell(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.NUMBER:
                case BIFFRECORDTYPE.NUMBER_OLD:
                    return new XlsBiffNumberCell(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.RK:
                    return new XlsBiffRKCell(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.MULRK:
                    return new XlsBiffMulRKCell(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.FORMULA:
                case BIFFRECORDTYPE.FORMULA_V3:
                case BIFFRECORDTYPE.FORMULA_V4:
                    return new XlsBiffFormulaCell(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.FORMAT_V23:
                case BIFFRECORDTYPE.FORMAT:
                    return new XlsBiffFormatString(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.STRING:
                case BIFFRECORDTYPE.STRING_OLD:
                    return new XlsBiffFormulaString(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.CONTINUE:
                    return new XlsBiffContinue(bytes, offset);
                case BIFFRECORDTYPE.DIMENSIONS:
                case BIFFRECORDTYPE.DIMENSIONS_V2:
                    return new XlsBiffDimensions(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.BOUNDSHEET:
                    return new XlsBiffBoundSheet(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.WINDOW1:
                    return new XlsBiffWindow1(bytes, offset);
                case BIFFRECORDTYPE.CODEPAGE:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.FNGROUPCOUNT:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.RECORD1904:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.BOOKBOOL:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.BACKUP:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.HIDEOBJ:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.USESELFS:
                    return new XlsBiffSimpleValueRecord(bytes, offset);
                case BIFFRECORDTYPE.UNCALCED:
                    return new XlsBiffUncalced(bytes, offset);
                case BIFFRECORDTYPE.QUICKTIP:
                    return new XlsBiffQuickTip(bytes, offset);
                case BIFFRECORDTYPE.MSODRAWING:
                    return new XlsBiffMSODrawing(bytes, offset);
                case BIFFRECORDTYPE.FILEPASS:
                    return new XlsBiffFilePass(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.HEADER:
                case BIFFRECORDTYPE.FOOTER:
                    return new XlsBiffHeaderFooterString(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.CODENAME:
                    return new XlsBiffCodeName(bytes, offset);
                case BIFFRECORDTYPE.XF:
                case BIFFRECORDTYPE.XF_V2:
                case BIFFRECORDTYPE.XF_V3:
                case BIFFRECORDTYPE.XF_V4:
                    return new XlsBiffXF(bytes, offset);
                case BIFFRECORDTYPE.MERGECELLS:
                    return new XlsBiffMergeCells(bytes, offset);
                case BIFFRECORDTYPE.COLINFO:
                    return new XlsBiffColInfo(bytes, offset);
                default:
                    return new XlsBiffRecord(bytes, offset);
            }
        }

        public void Dispose()
        {
            CipherTransform?.Dispose();
            ((IDisposable)Cipher)?.Dispose();
        }

        private int GetBiffVersion(XlsBiffBOF bof)
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
        /// Create an ICryptoTransform instance to decrypt a 1024-byte block
        /// </summary>
        private void CreateBlockDecryptor(int blockNumber)
        {
            CipherTransform?.Dispose();

            var blockKey = Encryption.GenerateBlockKey(blockNumber, SecretKey);
            CipherTransform = Cipher.CreateDecryptor(blockKey, null);
            CipherBlock = blockNumber;
        }

        /// <summary>
        /// Decrypt some dummy bytes to align the decryptor with the position in the current 1024-byte block
        /// </summary>
        private void AlignBlockDecryptor(int blockOffset)
        {
            var bytes = new byte[blockOffset];
            CryptoHelpers.DecryptBytes(CipherTransform, bytes);
        }

        private void DecryptRecord(int startPosition, BIFFRECORDTYPE id, byte[] bytes)
        {
            // Decrypt the last read record, find it's start offset relative to the current stream position
            int startDecrypt = 4;
            int recordSize = bytes.Length;
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
                var chunkSize = (int)Math.Min(recordSize - position, 1024 - blockOffset);
                var block = new byte[chunkSize];

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
}
