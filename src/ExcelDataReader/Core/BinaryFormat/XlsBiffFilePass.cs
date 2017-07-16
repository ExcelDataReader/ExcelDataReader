using System;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents FILEPASS record
    /// </summary>
    internal class XlsBiffFilePass : XlsBiffRecord
    {
        internal XlsBiffFilePass(byte[] bytes, uint offset)
            : base(bytes, offset)
        {
            ushort type = ReadUInt16(0);
            switch (type)
            {
                case 0:
                    throw new NotSupportedException("XOR obfuscation is not supported.");
                case 1:
                    break;
                default:
                    throw new NotSupportedException("Unknown encryption type: " + type);
            }

            ushort rc4Type = ReadUInt16(2);
            switch (rc4Type)
            {
                case 1:
                    break;
                case 2:
                case 3:
                    throw new NotSupportedException("CryptAPI is not supported.");
                default:
                    throw new NotSupportedException("Unknown RC4 encryption type: " + rc4Type);
            }
        }

        public byte[] Salt => ReadArray(6, 16);

        // Encryption info starts at byte 6.
        // If standard encryption:
        //   Two first bytes is 0x0001:  RC4 encryption header structure
        //   Two first bytes is 0x0002, 0x0003 or 0x0004: RC4 CryptoAPI encryption header structure
    }
}