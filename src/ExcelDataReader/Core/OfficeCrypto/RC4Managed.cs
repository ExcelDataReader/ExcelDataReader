using System;
using System.Security.Cryptography;

namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// Minimal RC4 decryption compatible with System.Security.Cryptography.SymmetricAlgorithm.
    /// </summary>
    internal class RC4Managed : SymmetricAlgorithm
    {
        public override ICryptoTransform CreateDecryptor(byte[] rgbKey, byte[] rgbIV)
        {
            return new RC4Transform(rgbKey);
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

        internal class RC4Transform : ICryptoTransform
        {
            private readonly byte[] _s = new byte[256];

            private int _index1;

            private int _index2;

            public RC4Transform(byte[] key)
            {
                Key = key;
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

            public int InputBlockSize => 1024;

            public int OutputBlockSize => 1024;

            public bool CanTransformMultipleBlocks => false;

            public bool CanReuseTransform => false;

            public byte[] Key { get; }

            public void Dispose()
            {
            }

            public int TransformBlock(byte[] inputBuffer, int inputOffset, int inputCount, byte[] outputBuffer, int outputOffset)
            {
                for (var i = 0; i < inputCount; i++)
                {
                    byte mask = Output();
                    outputBuffer[outputOffset + i] = (byte)(inputBuffer[inputOffset + i] ^ mask);
                }

                return inputCount;
            }

            public byte[] TransformFinalBlock(byte[] inputBuffer, int inputOffset, int inputCount)
            {
                var result = new byte[inputCount];
                TransformBlock(inputBuffer, inputOffset, inputCount, result, 0);
                return result;
            }

            private static void Swap(byte[] s, int i, int j)
            {
                byte c = s[i];

                s[i] = s[j];
                s[j] = c;
            }

            private byte Output()
            {
                _index1 = (_index1 + 1) & 255;
                _index2 = (_index2 + _s[_index1]) & 255;

                Swap(_s, _index1, _index2);

                return _s[(_s[_index1] + _s[_index2]) & 255];
            }
        }
    }
}
