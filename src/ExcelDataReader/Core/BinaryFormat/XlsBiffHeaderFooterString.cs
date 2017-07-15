using System;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a string value of a header or footer.
    /// </summary>
    internal sealed class XlsBiffHeaderFooterString : XlsBiffRecord
    {
        private readonly IXlsString _xlsString;

        internal XlsBiffHeaderFooterString(byte[] bytes, uint offset, int biffVersion)
            : base(bytes, offset)
        {
            if (biffVersion < 8)
                _xlsString = new XlsShortByteString(bytes, offset + 4);
            else if (biffVersion == 8)
                _xlsString = new XlsUnicodeString(bytes, offset + 4);
            else
                throw new ArgumentException("Unexpected BIFF version " + biffVersion, nameof(biffVersion));
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