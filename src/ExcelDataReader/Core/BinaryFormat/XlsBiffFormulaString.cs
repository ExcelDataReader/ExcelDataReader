using System.Text;

using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a string value of formula
    /// </summary>
    internal class XlsBiffFormulaString : XlsBiffRecord
    {
        private readonly XlsFormattedUnicodeString _unicodeString;

        internal XlsBiffFormulaString(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader)
        {
            _unicodeString = new XlsFormattedUnicodeString(bytes, offset + 4); 
        }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        public string Value => _unicodeString.Value;
    }
}