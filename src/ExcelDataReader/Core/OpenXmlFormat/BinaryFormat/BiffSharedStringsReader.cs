using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat
{
    internal sealed class BiffSharedStringsReader : BiffReader
    {
        private const int StringItem = 0x13;

        public BiffSharedStringsReader(Stream stream) 
            : base(stream)
        {
        }

        protected override Record ReadOverride(byte[] buffer, uint recordId, uint recordLength)
        {
            switch (recordId) 
            {
                case StringItem:
                    // Must be between 0 and 255 characters
                    uint length = GetDWord(buffer, 1);
                    string value = GetString(buffer, 1 + 4, length);
                    return new SharedStringRecord(value);
            }

            return Record.Default;
        }
    }
}
