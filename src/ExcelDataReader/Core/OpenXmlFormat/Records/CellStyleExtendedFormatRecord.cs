using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class CellStyleExtendedFormatRecord : Record
    {
        public CellStyleExtendedFormatRecord(ExtendedFormat extendedFormat)
        {
            ExtendedFormat = extendedFormat;
        }

        public ExtendedFormat ExtendedFormat { get; }
    }
}
