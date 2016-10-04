using System.Text;

namespace ExcelDataReader.Portable.Core.BinaryFormat
{
    internal interface IXlsString
    {
        /// <summary>
        /// Returns string represented by this instance
        /// </summary>
        string Value { get; }

        /// <summary>
        /// Count of characters in string
        /// </summary>
        ushort CharacterCount { get; }

        uint HeadSize { get;  }
        uint TailSize { get; }
        bool IsMultiByte { get; }
    }
}