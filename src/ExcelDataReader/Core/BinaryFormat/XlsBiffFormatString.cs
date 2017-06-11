using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a string value of format
    /// </summary>
    internal class XlsBiffFormatString : XlsBiffRecord
    {
        private readonly IXlsString _string;

        internal XlsBiffFormatString(byte[] bytes, uint offset, bool isV8, Encoding encoding)
            : base(bytes, offset)
        {
            if (isV8)
                _string = new XlsFormattedUnicodeString(bytes, offset + 6);
            else
                _string = new XlsByteString(bytes, offset + 4, encoding);
        }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        public string Value => _string.Value;

        public ushort Index
        {
            get
            {
                switch (Id)
                {
                    case BIFFRECORDTYPE.FORMAT_V23:
                        throw new NotSupportedException("Index is not available for BIFF2 and BIFF3 FORMAT records.");
                    default:
                        return ReadUInt16(0);
                }
            }
        }
    }
}