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
        private readonly ExcelBinaryReader _reader;
        private readonly byte[] _bytes;

        public XlsBiffStream(byte[] bytes, ExcelBinaryReader reader)
        {
            _reader = reader;
            _bytes = bytes;

            XlsBiffRecord rec = XlsBiffRecord.GetRecord(_bytes, 0, _reader);
            XlsBiffRecord rec2 = XlsBiffRecord.GetRecord(_bytes, (uint)rec.Size, reader);
            XlsBiffFilePass filePass = rec2 as XlsBiffFilePass;
            if (filePass == null)
            {
                XlsBiffRecord rec3 = XlsBiffRecord.GetRecord(_bytes, (uint)(rec.Size + rec2.Size), reader);
                filePass = rec3 as XlsBiffFilePass;
            }

            if (filePass != null)
            {
                RC4Key key = new RC4Key("VelvetSweatshop", filePass.Salt);

                int blockNumber = 0;
                RC4 rc4 = key.Create(blockNumber);

                int position = 0;
                while (position < _bytes.Length - 4)
                {
                    uint id = BitConverter.ToUInt16(_bytes, position);
                    int length = BitConverter.ToUInt16(_bytes, position + 2) + 4;

                    int startDecrypt = 4;
                    switch ((BIFFRECORDTYPE)id)
                    {
                        case BIFFRECORDTYPE.BOF:
                        case BIFFRECORDTYPE.FILEPASS:
                        case BIFFRECORDTYPE.INTERFACEHDR:
                            startDecrypt = length;
                            break;
                        case BIFFRECORDTYPE.BOUNDSHEET:
                            startDecrypt += 4; // For some reason the sheet offset is not encrypted
                            break;
                    }

                    for (int i = 0; i < length; i++)
                    {
                        int currentBlock = position / 1024;
                        if (blockNumber != currentBlock)
                        {
                            blockNumber = currentBlock;
                            rc4 = key.Create(blockNumber);
                        }

                        byte mask = rc4.Output();
                        if (i >= startDecrypt)
                        {
                            _bytes[position] = (byte)(_bytes[position] ^ mask);
                        }

                        position++;
                    }
                }
            }

            Size = _bytes.Length;
            Position = 0;
        }

        /// <summary>
        /// Gets the size of BIFF stream in bytes
        /// </summary>
        public int Size { get; }

        /// <summary>
        /// Gets the current position in BIFF stream
        /// </summary>
        public int Position { get; private set; }
        
        /// <summary>
        /// Sets stream pointer to the specified offset
        /// </summary>
        /// <param name="offset">Offset value</param>
        /// <param name="origin">Offset origin</param>
        public void Seek(int offset, SeekOrigin origin)
        {
            switch (origin)
            {
                case SeekOrigin.Begin:
                    Position = offset;
                    break;
                case SeekOrigin.Current:
                    Position += offset;
                    break;
                case SeekOrigin.End:
                    Position = Size - offset;
                    break;
            }

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
            if ((uint)Position + 4 >= _bytes.Length)
                return null;

            var record = XlsBiffRecord.GetRecord(_bytes, (uint)Position, _reader);

            if (record != null)
            {
                // Set readOption to loose to not cause exception here (sql reporting services)
                if (_reader.ReadOption == ReadOption.Strict)
                {
                    if (record.Bytes.Length < Position + record.Size)
                        throw new ArgumentException(Errors.ErrorBiffBufferSize);
                }

                Position += record.Size;
            }

            if (Position > Size)
            {
                record = null;
            }

            return record;
        }

        private sealed class RC4Key
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

        private sealed class RC4
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
