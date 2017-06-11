using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a string (max 255 bytes)
    /// </summary>
    internal class XlsBiffLabelCell : XlsBiffBlankCell
    {
        private readonly IXlsString _xlsString;

        internal XlsBiffLabelCell(byte[] bytes, uint offset, uint stringOffset, bool isV8, Encoding encoding)
            : base(bytes, offset)
        {
            _xlsString = XlsStringFactory.CreateXlsString(bytes, offset + stringOffset, isV8, encoding);
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