using System;
using System.IO;
using System.Security.Cryptography;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a BIFF stream
    /// </summary>
    internal class XlsBiffStream
    {
        public XlsBiffStream(Stream baseStream, int offset = 0, int explicitVersion = 0, RC4Key secretKey = null)
        {
            BaseStream = baseStream;
            Position = offset;

            var bof = Read() as XlsBiffBOF;
            if (bof != null)
            { 
                BiffVersion = explicitVersion == 0 ? GetBiffVersion(bof) : explicitVersion;
                BiffType = bof.Type;
            }

            if (secretKey != null)
            {
                SecretKey = secretKey;
            }
            else
            {
                var filePass = Read() as XlsBiffFilePass;
                if (filePass == null)
                    filePass = Read() as XlsBiffFilePass;

                if (filePass != null)
                    SecretKey = new RC4Key("VelvetSweatshop", filePass.Salt);
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

        public RC4Key SecretKey { get; }

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
                    return new XlsBiffFilePass(bytes, offset);
                case BIFFRECORDTYPE.HEADER:
                case BIFFRECORDTYPE.FOOTER:
                    return new XlsBiffHeaderFooterString(bytes, offset, biffVersion);
                case BIFFRECORDTYPE.CODENAME:
                    return new XlsBiffCodeName(bytes, offset);

                default:
                    return new XlsBiffRecord(bytes, offset);
            }
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
                    if (bof.Version == 0x500)
                        return 5;
                    if (bof.Version == 0x600)
                        return 8;
                    break;
            }

            return 0;
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
                RC4 rc4 = SecretKey.Create(blockNumber);

                for (var i = 0; i < blockOffset; i++)
                    rc4.Output();

                var chunkSize = (int)Math.Min(recordSize - position, 1024 - blockOffset);
                for (var i = 0; i < chunkSize; i++)
                {
                    byte mask = rc4.Output();
                    if (position >= startDecrypt)
                        bytes[position] ^= mask;

                    position++;
                }
            }
        }

        internal sealed class RC4Key
        {
            private readonly byte[] _key;

            public RC4Key(string password, byte[] salt)
            {
                int length = Math.Min(password.Length, 16);
                byte[] passwordData = new byte[length * 2];
                for (int i = 0; i < length; i++)
                {
                    char ch = password[i];
                    passwordData[i * 2 + 0] = (byte)((ch << 0) & 0xFF);
                    passwordData[i * 2 + 1] = (byte)((ch << 8) & 0xFF);
                }

                using (MD5 md5 = MD5.Create())
                {
                    byte[] passwordHash = md5.ComputeHash(passwordData);

                    md5.Initialize();

                    const int truncateCount = 5;
                    byte[] intermediateData = new byte[truncateCount * 16 + salt.Length * 16];

                    int offset = 0;
                    for (int i = 0; i < 16; i++)
                    {
                        Array.Copy(passwordHash, 0, intermediateData, offset, truncateCount);
                        offset += truncateCount;
                        Array.Copy(salt, 0, intermediateData, offset, salt.Length);
                        offset += salt.Length;
                    }

                    const int keyLength = 5;

                    byte[] finalHash = md5.ComputeHash(intermediateData);
                    byte[] result = new byte[keyLength];
                    Array.Copy(finalHash, 0, result, 0, keyLength);

                    _key = result;
                }
            }
            
            public RC4 Create(int blockNumber)
            {
                byte[] data = new byte[4 + _key.Length];
                data[data.Length - 1] = (byte)((blockNumber >> 24) & 0xFF);
                data[data.Length - 2] = (byte)((blockNumber >> 16) & 0xFF);
                data[data.Length - 3] = (byte)((blockNumber >> 8) & 0xFF);
                data[data.Length - 4] = (byte)((blockNumber >> 0) & 0xFF);

                Array.Copy(_key, 0, data, 0, _key.Length);

                using (MD5 md5 = MD5.Create())
                {
                    byte[] blockKey = md5.ComputeHash(data);

                    return new RC4(blockKey);
                }
            }
        }

        internal sealed class RC4
        {
            private readonly byte[] _s = new byte[256];

            private int _index1;

            private int _index2;

            public RC4(byte[] key)
            {
                for (int i = 0; i < _s.Length; i++)
                {
                    _s[i] = (byte)i;
                }

                for (int i = 0, j = 0; i < 256; i++)
                {
                    j = (j + key[i % key.Length] + _s[i]) & 255;

                    Swap(_s, i, j);
                }
            }

            public byte Output()
            {
                _index1 = (_index1 + 1) & 255;
                _index2 = (_index2 + _s[_index1]) & 255;

                Swap(_s, _index1, _index2);

                return _s[(_s[_index1] + _s[_index2]) & 255];
            }

            private static void Swap(byte[] s, int i, int j)
            {
                byte c = s[i];

                s[i] = s[j];
                s[j] = c;
            }
        }
    }
}
