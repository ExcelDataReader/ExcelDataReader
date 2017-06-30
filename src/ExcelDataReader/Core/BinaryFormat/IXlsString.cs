using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    internal interface IXlsString
    {
        /// <summary>
        /// Gets the nubmer of characters in the string.
        /// </summary>
        ushort CharacterCount { get; }

        uint HeadSize { get;  }

        uint TailSize { get; }

        bool IsMultiByte { get; }

        /// <summary>
        /// Gets the string value. Encoding is only used with BIFF2-5 byte strings.
        /// </summary>
        string GetValue(Encoding encoding);
    }
}