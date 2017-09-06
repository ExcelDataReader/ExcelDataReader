using System;
using System.IO;

namespace ExcelDataReader.Core.OfficeCrypto
{
    /// <summary>
    /// A seekable stream for reading an EncryptedPackage blob using OpenXml Agile Encryption. 
    /// </summary>
    internal class AgileEncryptedPackageStream : Stream
    {
        private const int SegmentLength = 4096;

        public AgileEncryptedPackageStream(Stream stream, byte[] key, byte[] iv, EncryptionInfo encryption)
        {
            Stream = stream;
            Key = key;
            IV = iv;
            Encryption = encryption;

            Stream.Read(SegmentBytes, 0, 8);
            DecryptedLength = BitConverter.ToInt32(SegmentBytes, 0);
            ReadSegment();
        }

        public override bool CanRead => true;

        public override bool CanSeek => true;

        public override bool CanWrite => false;

        public override long Length => DecryptedLength;

        public override long Position { get => Offset - SegmentLength + SegmentOffset; set => Seek(value, SeekOrigin.Begin); }

        private Stream Stream { get; set; }

        private byte[] Key { get; }

        private byte[] IV { get; }

        private HashIdentifier HashAlgorithm { get; }

        private EncryptionInfo Encryption { get; }

        private int Offset { get; set; }

        private byte[] SegmentBytes { get; set; } = new byte[SegmentLength];

        private int SegmentOffset { get; set; }

        private int SegmentIndex { get; set; }

        private int DecryptedLength { get; set; }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (Position >= Length)
            {
                throw new InvalidOperationException("Tried to read past the end of the encrypted stream");
            }

            int index = 0;
            while (index < count)
            {
                if (SegmentOffset == SegmentBytes.Length)
                {
                    ReadSegment();
                    SegmentOffset = 0;
                }

                var chunkSize = Math.Min(count - index, SegmentBytes.Length - SegmentOffset);
                Array.Copy(SegmentBytes, SegmentOffset, buffer, offset + index, chunkSize);
                index += chunkSize;
                SegmentOffset += chunkSize;
            }

            return index;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            switch (origin)
            {
                case SeekOrigin.Begin:
                    SegmentIndex = (int)(offset / SegmentLength);
                    Offset = SegmentIndex * SegmentLength;
                    SegmentOffset = (int)(offset % SegmentLength);
                    if (Offset < Length)
                        ReadSegment();
                    return Position;
                case SeekOrigin.Current:
                    return Seek(Position + offset, SeekOrigin.Begin);
                case SeekOrigin.End:
                    return Seek(Length + offset, SeekOrigin.Begin);
                default:
                    return Offset;
            }
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Stream?.Dispose();
                Stream = null;
            }

            base.Dispose(disposing);
        }

        private void ReadSegment()
        {
            var salt = Encryption.GenerateBlockKey(SegmentIndex, IV);
            
            // NOTE: +8 skips EncryptedPackage header
            Stream.Seek(8 + Offset, SeekOrigin.Begin);
            Stream.Read(SegmentBytes, 0, SegmentLength);

            using (var cipher = Encryption.CreateCipher())
            {
                SegmentBytes = CryptoHelpers.DecryptBytes(cipher, SegmentBytes, Key, salt);
            }

            SegmentIndex++;
            Offset += SegmentLength;
        }
    }
}
