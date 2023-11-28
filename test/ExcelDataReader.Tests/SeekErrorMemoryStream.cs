using System;
using System.IO;

namespace ExcelDataReader.Tests
{
    public class SeekErrorMemoryStream : MemoryStream
    {
        private bool _canSeek;

        private SeekErrorMemoryStream()
        {
        }

        public override bool CanSeek => _canSeek;

        /// <summary>
        /// Creates SeekErrorMemoryStream copy data from the source
        /// </summary>
        public static SeekErrorMemoryStream CreateFromStream(Stream source)
        {
            var forwardStream = new SeekErrorMemoryStream { _canSeek = true };

            source.CopyTo(forwardStream);

            forwardStream.Seek(0, SeekOrigin.Begin);

            // now disable seek
            forwardStream._canSeek = false;

            return forwardStream;
        }

        public override long Seek(long offset, SeekOrigin loc)
        {
            if (_canSeek)
                return base.Seek(offset, loc);

            // throw offset error to simulate problem we had with HttpInputStream
            throw new ArgumentOutOfRangeException(nameof(offset));
        }
    }
}
