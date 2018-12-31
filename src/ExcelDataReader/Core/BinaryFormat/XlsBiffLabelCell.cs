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

        internal XlsBiffLabelCell(byte[] bytes, uint offset, int biffVersion)
            : base(bytes, offset, biffVersion)
        {
            if (Id == BIFFRECORDTYPE.LABEL_OLD)
            {
                // BIFF2
                _xlsString = new XlsShortByteString(bytes, offset + 4 + 7);
            }
            else if (biffVersion >= 2 && biffVersion <= 5)
            {
                // BIFF3-5, or if there is a newer label record present in a BIFF2 stream
                _xlsString = new XlsByteString(bytes, offset + 4 + 6);
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
        /// Gets the cell value.
        /// </summary>
        public string GetValue(Encoding encoding)
        {
            return _xlsString.GetValue(encoding);
        }
    }
}