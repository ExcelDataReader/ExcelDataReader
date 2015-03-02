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

        Encoding UseEncoding { //get { return IsMultiByte ? Encoding.Unicode : Encoding.UTF8; } 
            //not sure this is a good assumption but it does work for every test case
            get; }
    }
}