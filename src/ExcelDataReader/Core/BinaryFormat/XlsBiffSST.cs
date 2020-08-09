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
        private readonly XlsSSTReader _reader = new XlsSSTReader();

        internal XlsBiffSST(byte[] bytes)
            : base(bytes)
        {
            _strings = new List<IXlsString>();
            ReadSstStrings();
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
        /// Parses strings out of a Continue record
        /// </summary>
        public void ReadContinueStrings(XlsBiffContinue sstContinue)
        {
            if (_strings.Count == UniqueCount)
            {
                return;
            }

            foreach (var str in _reader.ReadStringsFromContinue(sstContinue))
            {
                _strings.Add(str);

                if (_strings.Count == UniqueCount)
                {
                    break;
                }
            }
        }

        public void Flush()
        {
            var str = _reader.Flush();
            if (str != null)
            {
                _strings.Add(str);
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

            return null; // #VALUE error
        }

        /// <summary>
        /// Parses strings out of this SST record
        /// </summary>
        private void ReadSstStrings()
        {
            if (_strings.Count == UniqueCount)
            {
                return;
            }

            foreach (var str in _reader.ReadStringsFromSST(this))
            {
                _strings.Add(str);

                if (_strings.Count == UniqueCount)
                {
                    break;
                }
            }
        }
    }
}