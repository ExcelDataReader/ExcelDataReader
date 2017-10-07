using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelDataReader.Core.CsvFormat
{
    internal class CsvWorkbook : IWorkbook<CsvWorksheet>
    {
        public CsvWorkbook(Stream stream, Encoding encoding)
        {
            Stream = stream;
            Encoding = encoding;
        }

        public int ResultsCount => 1;

        public Stream Stream { get; }

        public Encoding Encoding { get; }

        public IEnumerable<CsvWorksheet> ReadWorksheets()
        {
            yield return new CsvWorksheet(Stream, Encoding);
        }
    }
}
