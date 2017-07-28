using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a Shared String Table in BIFF8 format
    /// </summary>
    internal class XlsBiffSST : XlsBiffRecord
    {
        private readonly List<IXlsString> _strings;

        internal XlsBiffSST(byte[] bytes, uint offset)
            : base(bytes, offset)
        {
            _strings = new List<IXlsString>();
        }

        /// <summary>
        /// Gets the number of strings in SST
        /// </summary>
        public uint Count => ReadUInt32(0x0);

        /// <summary>
        /// Gets the count of unique strings in SST
        /// </summary>
        public uint UniqueCount => ReadUInt32(0x4);

        /// <summary>
        /// Parses strings out of the SST record and subsequent Continue records from the BIFF stream
        /// </summary>
        public void ReadStrings(XlsBiffStream biffStream)
        {
            var reader = new XlsSSTReader(this, biffStream);

            for (var i = 0; i < UniqueCount; i++)
            {
                var s = reader.ReadString();
                _strings.Add(s);
            }
        }
        
        /// <summary>
        /// Returns string at specified index
        /// </summary>
        /// <param name="sstIndex">Index of string to get</param>
        /// <param name="encoding">Workbook encoding</param>
        /// <returns>string value if it was found, empty string otherwise</returns>
        public string GetString(uint sstIndex, Encoding encoding)
        {
            if (sstIndex < _strings.Count)
                return _strings[(int)sstIndex].GetValue(encoding);
            
            return string.Empty;
        }
    }
}