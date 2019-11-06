using System;
using ExcelDataReader.Core.OfficeCrypto;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents FILEPASS record containing XOR obfuscation details or a an EncryptionInfo structure
    /// </summary>
    internal class XlsBiffFilePass : XlsBiffRecord
    {
        internal XlsBiffFilePass(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            if (biffVersion >= 2 && biffVersion <= 5)
            {
                // Cipher = EncryptionType.XOR;
                var encryptionKey = ReadUInt16(0);
                var hashValue = ReadUInt16(2);
                EncryptionInfo = EncryptionInfo.Create(encryptionKey, hashValue);
            }
            else
            {
                ushort type = ReadUInt16(0);

                if (type == 0)
                {
                    var encryptionKey = ReadUInt16(2);
                    var hashValue = ReadUInt16(4);
                    EncryptionInfo = EncryptionInfo.Create(encryptionKey, hashValue);
                }
                else if (type == 1)
                {
                    var encryptionInfo = new byte[bytes.Length - 6]; // 6 = 4 + 2 = biffVersion header + filepass enryptiontype
                    Array.Copy(bytes, 6, encryptionInfo, 0, bytes.Length - 6);
                    EncryptionInfo = EncryptionInfo.Create(encryptionInfo);
                }
                else
                {
                    throw new NotSupportedException("Unknown encryption type: " + type);
                }
            }
        }

        public EncryptionInfo EncryptionInfo { get; }
    }
}