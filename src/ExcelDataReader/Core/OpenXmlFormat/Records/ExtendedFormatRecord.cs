using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class ExtendedFormatRecord : Record
    {
        public ExtendedFormatRecord(int id, int formatIndexInFile, bool applyNumberFormat) 
        {
            Id = id;
            FormatIndexInFile = formatIndexInFile;
            ApplyNumberFormat = applyNumberFormat;
        }

        public int Id { get; }

        public int FormatIndexInFile { get; }

        public bool ApplyNumberFormat { get; }
    }
}
