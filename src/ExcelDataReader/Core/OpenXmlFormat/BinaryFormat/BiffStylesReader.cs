using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat
{
    internal sealed class BiffStylesReader : BiffReader
    {
        private const int CellXfStart = 0x269;
        private const int Xf = 0x2f;
        private const int CellXfEnd = 0x26a;

        private const int NumberFormatStart = 0x267;
        private const int NumberFormat = 0x2c;
        private const int NumberFormatEnd = 0x268;

        private bool _inCellXf;
        private bool _inNumberFormat;

        public BiffStylesReader(Stream stream)
            : base(stream)
        {
        }

        protected override Record ReadOverride(byte[] buffer, uint recordId, uint recordLength)
        {
            switch (recordId)
            {
                case CellXfStart:
                    _inCellXf = true;
                    break;
                case CellXfEnd:
                    _inCellXf = false;
                    break;
                case NumberFormatStart:
                    _inNumberFormat = true;
                    break;
                case NumberFormatEnd:
                    _inNumberFormat = false;
                    break;

                case Xf when _inCellXf:
                    {
                        int id = GetWord(buffer, 0);
                        int format = GetWord(buffer, 2);
                        bool applyNumberFormat = (buffer[14] & 1) != 0;

                        return new ExtendedFormatRecord(id, format, applyNumberFormat);
                    }

                case NumberFormat when _inNumberFormat:
                    {
                        // Must be between 1 and 255 characters
                        int format = GetWord(buffer, 0);
                        uint length = GetDWord(buffer, 2);
                        string formatString = GetString(buffer, 2 + 4, length);

                        return new NumberFormatRecord(format, formatString);
                    }
            }

            return Record.Default;
        }
    }
}
