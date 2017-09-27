using System;
using System.Security.Cryptography;

namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Minimal Office "XOR Deobfuscation Method 1" implementation compatible
    /// with System.Security.Cryptography.SymmetricAlgorithm.
    /// </summary>
    internal class XorManaged : SymmetricAlgorithm
    {
        private static byte[] padArray = new byte[]
        {
            0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80,
            0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00
        };

        private static ushort[] initialCode = new ushort[]
        {
            0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C,
            0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139,
            0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3
        };

        private static ushort[] xorMatrix = new ushort[]
        {
            0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09,
            0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF,
            0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0,
            0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40,
            0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5,
            0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A,
            0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9,
            0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0,
            0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC,
            0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10,
            0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168,
            0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C,
            0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD,
            0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC,
            0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4
        };

        public XorManaged()
        {
        }

        public override ICryptoTransform CreateDecryptor(byte[] rgbKey, byte[] rgbIV)
        {
            return new XorTransform(rgbKey, 0);
        }

        public override ICryptoTransform CreateEncryptor(byte[] rgbKey, byte[] rgbIV)
        {
            throw new NotImplementedException();
        }

        public override void GenerateIV()
        {
            throw new NotImplementedException();
        }

        public override void GenerateKey()
        {
            throw new NotImplementedException();
        }

        internal static ushort CreatePasswordVerifier_Method1(byte[] passwordBytes)
        {
            var passwordArray = CryptoHelpers.Combine(new byte[] { (byte)passwordBytes.Length }, passwordBytes);
            ushort verifier = 0x0000;
            for (var i = 0; i < passwordArray.Length; ++i)
            {
                var passwordByte = passwordArray[passwordArray.Length - 1 - i];
                ushort intermediate1 = (ushort)(((verifier & 0x4000) == 0) ? 0 : 1);
                ushort intermediate2 = (ushort)(verifier * 2);
                intermediate2 &= 0x7FFF;
                ushort intermediate3 = (ushort)(intermediate1 | intermediate2);
                verifier = (ushort)(intermediate3 ^ passwordByte);
            }

            return (ushort)(verifier ^ 0xCE4B);
        }

        internal static ushort CreateXorKey_Method1(byte[] passwordBytes)
        {
            ushort xorKey = initialCode[passwordBytes.Length - 1];
            var currentElement = 0x68;

            for (var i = 0; i < passwordBytes.Length; ++i)
            {
                var c = passwordBytes[passwordBytes.Length - 1 - i];
                for (var j = 0; j < 7; ++j)
                {
                    if ((c & 0x40) != 0)
                    {
                        xorKey ^= xorMatrix[currentElement];
                    }

                    c *= 2;
                    currentElement--;
                }
            }

            return xorKey;
        }

        /// <summary>
        /// Generates a 16 byte obfuscation array based on the POI/LibreOffice implementations
        /// </summary>
        internal static byte[] CreateXorArray_Method1(byte[] passwordBytes)
        {
            var index = passwordBytes.Length;
            var obfuscationArray = new byte[16];
            Array.Copy(passwordBytes, 0, obfuscationArray, 0, passwordBytes.Length);
            Array.Copy(padArray, 0, obfuscationArray, passwordBytes.Length, padArray.Length - passwordBytes.Length + 1);

            var xorKey = CreateXorKey_Method1(passwordBytes);
            byte[] baseKeyLE = new byte[] { (byte)(xorKey & 0xFF), (byte)((xorKey >> 8) & 0xFF) };
            int nRotateSize = 2;
            for (int i = 0; i < obfuscationArray.Length; i++)
            {
                obfuscationArray[i] ^= baseKeyLE[i & 1];
                obfuscationArray[i] = RotateLeft(obfuscationArray[i], nRotateSize);
            }

            return obfuscationArray;
        }

        private static byte RotateLeft(byte b, int shift)
        {
            return (byte)(((b << shift) | (b >> (8 - shift))) & 0xFF);
        }

        internal class XorTransform : ICryptoTransform
        {
            public XorTransform(byte[] key, int xorArrayIndex)
            {
                XorArray = key;
                XorArrayIndex = xorArrayIndex;
            }

            public int InputBlockSize => 1024;

            public int OutputBlockSize => 1024;

            public bool CanTransformMultipleBlocks => false;

            public bool CanReuseTransform => false;

            /// <summary>
            /// Gets or sets the obfuscation array index. BIFF obfuscation uses a different XorArrayIndex per record.
            /// </summary>
            public int XorArrayIndex { get; set; }

            private byte[] XorArray { get; }

            public void Dispose()
            {
            }

            public int TransformBlock(byte[] inputBuffer, int inputOffset, int inputCount, byte[] outputBuffer, int outputOffset)
            {
                for (var i = 0; i < inputCount; ++i)
                {
                    var value = inputBuffer[inputOffset + i];
                    value = RotateLeft(value, 3);
                    value ^= XorArray[XorArrayIndex % 16];
                    outputBuffer[outputOffset + i] = value;
                    XorArrayIndex++;
                }

                return inputCount;
            }

            public byte[] TransformFinalBlock(byte[] inputBuffer, int inputOffset, int inputCount)
            {
                var result = new byte[inputCount];
                TransformBlock(inputBuffer, inputOffset, inputCount, result, 0);
                return result;
            }
        }
    }
}
