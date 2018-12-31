using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a string value of format
    /// </summary>
    internal class XlsBiffFormatString : XlsBiffRecord
    {
        private readonly IXlsString _xlsString;

        internal XlsBiffFormatString(byte[] bytes, uint offset, int biffVersion)
            : base(bytes, offset)
        {
            if (Id == BIFFRECORDTYPE.FORMAT_V23)
            {
                // BIFF2-3
                _xlsString = new XlsShortByteString(bytes, offset + 4);
            }
            else if (biffVersion >= 2 && biffVersion <= 5)
            {
                // BIFF4-5, or if there is a newer format record in a BIFF2-3 stream
                _xlsString = new XlsShortByteString(bytes, offset + 4 + 2);
            }
            else if (biffVersion == 8)
            {
                // BIFF8
                _xlsString = new XlsUnicodeString(bytes, offset + 4 + 2);
            }
            else
            {
                throw new ArgumentException("Unexpected BIFF version " + biffVersion.ToString(), nameof(biffVersion));
            }
        }

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

        /// <summary>
        /// Gets the string value.
        /// </summary>
        public string GetValue(Encoding encoding)
        {
            return _xlsString.GetValue(encoding);
        }
    }
}