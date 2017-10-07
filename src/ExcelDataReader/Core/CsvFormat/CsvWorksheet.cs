using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelDataReader.Core.CsvFormat
{
    internal class CsvWorksheet : IWorksheet
    {
        private const char Separator = ',';

        public CsvWorksheet(Stream stream, Encoding fallbackEncoding)
        {
            Stream = stream;
            Stream.Seek(0, SeekOrigin.Begin);
            try
            {
                // Try as UTF-8 first, or use BOM if present
                var reader = new CsvReader(Stream, Separator, Encoding.UTF8);
                FieldCount = ReadFieldCount(reader);
                Encoding = reader.Encoding;
            }
            catch (DecoderFallbackException)
            {
                // If cannot parse as UTF-8, try fallback encoding
                Stream.Seek(0, SeekOrigin.Begin);

                var reader = new CsvReader(Stream, Separator, fallbackEncoding);
                FieldCount = ReadFieldCount(reader);
                Encoding = reader.Encoding;
            }
        }

        public string Name => string.Empty;

        public string CodeName => null;

        public string VisibleState => null;

        public HeaderFooter HeaderFooter => null;

        public int FieldCount { get; }

        public Stream Stream { get; }

        public Encoding Encoding { get; }

        public IEnumerable<Row> ReadRows()
        {
            Stream.Seek(0, SeekOrigin.Begin);

            var reader = new CsvReader(Stream, Separator, Encoding);

            var rowIndex = 0;
            while (true)
            {
                var row = reader.ReadRow();
                if (row == null)
                {
                    break;
                }

                var columnIndex = 0;

                var cells = new List<Cell>();
                foreach (var item in row)
                {
                    cells.Add(new Cell()
                    {
                        ColumnIndex = columnIndex,
                        Value = item
                    });

                    columnIndex++;
                }

                yield return new Row()
                {
                    Height = 12.75, // 255 twips
                    Cells = cells,
                    RowIndex = rowIndex
                };
            }
        }

        private int ReadFieldCount(CsvReader reader)
        {
            var result = 0;
            while (true)
            {
                var row = reader.ReadRow();
                if (row == null)
                {
                    break;
                }

                result = Math.Max(result, row.Count);
            }

            return result;
        }
    }
}
