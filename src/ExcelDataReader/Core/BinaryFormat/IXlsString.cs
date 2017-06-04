namespace ExcelDataReader.Core.BinaryFormat
{
    internal interface IXlsString
    {
        /// <summary>
        /// Gets the string value.
        /// </summary>
        string Value { get; }

        /// <summary>
        /// Gets the nubmer of characters in the string.
        /// </summary>
        ushort CharacterCount { get; }

        uint HeadSize { get;  }

        uint TailSize { get; }

        bool IsMultiByte { get; }
    }
}