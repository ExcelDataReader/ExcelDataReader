using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// [MS-XLS] 2.4.148 Label
    /// Represents a string
    /// </summary>
    internal class XlsBiffLabelCell : XlsBiffBlankCell
    {
        private readonly IXlsString _xlsString;

        internal XlsBiffLabelCell(byte[] bytes, uint offset, int biffVersion, Encoding encoding)
            : base(bytes, offset)
        {
            if (biffVersion == 2)
            {
                // BIFF2
                _xlsString = new XlsShortByteString(bytes, offset + 4 + 7, encoding);
            }
            else if (biffVersion >= 3 && biffVersion <= 5)
            {
                // BIFF3-5
                _xlsString = new XlsByteString(bytes, offset + 4 + 6, encoding);
            }
            else if (biffVersion == 8)
            {
                // BIFF8
                _xlsString = new XlsUnicodeString(bytes, offset + 4 + 6);
            }
            else
            {
                throw new ArgumentException("Unexpected BIFF version " + biffVersion.ToString(), nameof(biffVersion));
            }
        }

        /// <summary>
        /// Gets the length of string value
        /// </summary>
        public ushort Length => _xlsString.CharacterCount;

        /// <summary>
        /// Gets the cell value.
        /// </summary>
        public string Value => _xlsString.Value;
    }
}