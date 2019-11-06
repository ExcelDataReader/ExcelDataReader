using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a string value of formula
    /// </summary>
    internal class XlsBiffFormulaString : XlsBiffRecord
    {
        private readonly IXlsString _xlsString;

        internal XlsBiffFormulaString(byte[] bytes, int biffVersion)
            : base(bytes)
        {
            if (biffVersion == 2)
            {
                // BIFF2
                _xlsString = new XlsShortByteString(bytes, ContentOffset);
            }
            else if (biffVersion >= 3 && biffVersion <= 5)
            {
                // BIFF3-5
                _xlsString = new XlsByteString(bytes, ContentOffset);
            }
            else if (biffVersion == 8)
            {
                // BIFF8
                _xlsString = new XlsUnicodeString(bytes, ContentOffset);
            }
            else
            {
                throw new ArgumentException("Unexpected BIFF version " + biffVersion, nameof(biffVersion));
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